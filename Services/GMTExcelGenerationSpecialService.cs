using ClosedXML.Excel;
using GMTReportGenerater_api.Models;

namespace GMTReportGenerater_api.Services;

/// <summary>
/// GMT报告Excel文件生成服务 - Special格式（使用ClosedXML - 完全免费）
/// </summary>
public class GMTExcelGenerationSpecialService
{
    private readonly string _templatePath;
    private readonly string _outputDirectory;

    public GMTExcelGenerationSpecialService(IWebHostEnvironment environment)
    {
        _templatePath = Path.Combine(environment.ContentRootPath, "Template", "GMT-BB30-SPECIAL_template.xlsx");
        _outputDirectory = Path.Combine(environment.ContentRootPath, "Output");
        
        // 尝试创建输出目录，如果失败则记录日志但不抛出异常
        try
        {
            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine($"[警告] 无法创建Output文件夹，请手动创建并设置权限: {_outputDirectory}");
            Console.WriteLine($"[错误详情] {ex.Message}");
            // 不抛出异常，让应用继续启动，在生成文件时再处理
        }
    }

    /// <summary>
    /// 生成GMT月报Excel文件 - Special格式
    /// </summary>
    public async Task<string> GenerateExcelReportAsync(GMTReportResponse response, DateTime startDate, DateTime endDate)
    {
        // 检查模板文件是否存在
        if (!File.Exists(_templatePath))
        {
            throw new FileNotFoundException($"Special模板文件不存在: {_templatePath}");
        }

        // 生成输出文件名
        var fileName = $"GMT_Report_Special_{startDate:yyyyMMdd}_{endDate:yyyyMMdd}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        var outputPath = Path.Combine(_outputDirectory, fileName);

        // 使用ClosedXML打开模板并填充数据
        using (var workbook = new XLWorkbook(_templatePath))
        {
            foreach (var deviceSummary in response.DeviceSummaries)
            {
                foreach (var cpData in deviceSummary.CPData)
                {
                    // 提取DEVICE前6位
                    var baseDeviceName = deviceSummary.DeviceName.Length >= 6 
                        ? deviceSummary.DeviceName.Substring(0, 6) 
                        : deviceSummary.DeviceName;
                    
                    // 根据DEVICE前6位和CP号确定工作表名
                    var sheetName = $"{baseDeviceName}_{cpData.CPNumber}";
                    
                    Console.WriteLine($"[INFO] Special格式 - 查找工作表: {sheetName}");
                    
                    // 查找模板中的工作表
                    if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet))
                    {
                        Console.WriteLine($"[WARNING] Special模板中未找到工作表 {sheetName}，跳过该设备");
                        continue;
                    }
                    
                    Console.WriteLine($"[INFO] 找到工作表 {sheetName}，开始填充Special格式数据");

                    // 填充数据（从第3行开始，Special格式）
                    int row = 3;
                    
                    // 按Device+WF_LOT分组（Special格式需要完整的Device信息）
                    var deviceLotGroups = cpData.DataRows
                        .GroupBy(r => new { Device = deviceSummary.DeviceName, r.WF_Lot })
                        .OrderBy(g => g.Key.Device)
                        .ThenBy(g => g.Key.WF_Lot);
                    
                    foreach (var group in deviceLotGroups)
                    {
                        var groupRows = group.OrderBy(r => int.Parse(r.WF_No)).ToList();
                        
                        // 填充该组的所有Wafer数据
                        foreach (var dataRow in groupRows)
                        {
                            int col = 1;
                            // 日期格式转换为YYYY/MM/DD
                            var formattedDate = FormatDateToYYYYMMDD(dataRow.Date);
                            worksheet.Cell(row, col++).Value = formattedDate;                   // Date
                            worksheet.Cell(row, col++).Value = "HTSH";                          // Location
                            worksheet.Cell(row, col++).Value = group.Key.Device;                // Device (完整设备名)
                            worksheet.Cell(row, col++).Value = cpData.CPNumber;                 // CP_NO
                            worksheet.Cell(row, col++).Value = dataRow.WF_Lot;                  // WF_LOT
                            worksheet.Cell(row, col++).Value = dataRow.WF_No;                   // WF_NO
                            
                            // 计算Pass和Fail数量
                            int passQty = 0;
                            int totalQty = 0;
                            var cpConfig = GetCPConfigFromDataRow(dataRow);
                            
                            foreach (var bin in dataRow.BinCounts)
                            {
                                totalQty += bin.Value;
                                // 假设BIN键包含在PassBins中则为Pass
                                if (dataRow.Yield != "0.0%" && float.Parse(dataRow.Yield.Replace("%", "")) > 0)
                                {
                                    // 通过Yield反推Pass数量
                                    var yieldValue = float.Parse(dataRow.Yield.Replace("%", "")) / 100;
                                    passQty = (int)(totalQty * yieldValue);
                                }
                            }
                            
                            int failQty = totalQty - passQty;
                            
                            worksheet.Cell(row, col++).Value = totalQty;                        // GROSS_QTY
                            worksheet.Cell(row, col++).Value = passQty;                         // PASS_QTY
                            worksheet.Cell(row, col++).Value = failQty;                         // FAIL_QTY
                            worksheet.Cell(row, col++).Value = dataRow.Yield;                   // YIELD
                            
                            // 填充BIN数据
                            foreach (var bin in dataRow.BinCounts.OrderBy(b => b.Key))
                            {
                                worksheet.Cell(row, col++).Value = bin.Value;
                            }
                            
                            // Tester 和 Card_ID
                            worksheet.Cell(row, col++).Value = dataRow.Tester;                  // TESTER
                            worksheet.Cell(row, col++).Value = dataRow.Card_ID;                 // CARD_ID
                            
                            row++;
                        }
                    }
                    
                    // 设置整个数据区域的格式
                    if (cpData.DataRows.Any())
                    {
                        int lastDataCol = 12 + cpData.DataRows.First().BinCounts.Count; // Special格式列更多
                        int lastRow = row - 1;
                        int firstRow = 3;
                        
                        // 为每个单元格设置边框和黄色背景
                        for (int r = firstRow; r <= lastRow; r++)
                        {
                            for (int c = 1; c <= lastDataCol; c++)
                            {
                                var cell = worksheet.Cell(r, c);
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                        }
                    }
                }
            }

            // 保存文件
            await Task.Run(() => workbook.SaveAs(outputPath));
        }

        Console.WriteLine($"[INFO] Special格式Excel文件已生成: {outputPath}");
        return outputPath;
    }

    private dynamic GetCPConfigFromDataRow(ProcessedDataRow dataRow)
    {
        // 简化版本，实际应该从ConfigService获取
        return new { PassBins = new List<string>() };
    }

    /// <summary>
    /// 将列号转换为Excel列字母
    /// </summary>
    private string GetColumnLetter(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }

    /// <summary>
    /// 将日期字符串格式化为YYYY/MM/DD
    /// </summary>
    private string FormatDateToYYYYMMDD(string dateStr)
    {
        if (string.IsNullOrEmpty(dateStr))
            return string.Empty;

        try
        {
            // 尝试解析日期字符串（支持多种格式）
            if (DateTime.TryParse(dateStr, out DateTime date))
            {
                return date.ToString("yyyy/MM/dd");
            }

            // 如果已经是YYYY-MM-DD格式，直接替换
            if (dateStr.Length >= 10 && dateStr.Contains("-"))
            {
                return dateStr.Substring(0, 10).Replace("-", "/");
            }

            return dateStr;
        }
        catch
        {
            return dateStr;
        }
    }
}
