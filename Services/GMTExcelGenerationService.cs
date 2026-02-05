using ClosedXML.Excel;
using GMTReportGenerater_api.Models;

namespace GMTReportGenerater_api.Services;

/// <summary>
/// GMT报告Excel文件生成服务（使用ClosedXML - 完全免费）
/// </summary>
public class GMTExcelGenerationService
{
    private readonly string _templatePath;
    private readonly string _outputDirectory;

    public GMTExcelGenerationService(IWebHostEnvironment environment)
    {
        _templatePath = Path.Combine(environment.ContentRootPath, "Template", "GMT-BB30_template.xlsx");
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
    /// 生成GMT月报Excel文件
    /// </summary>
    /// <param name="response">处理后的GMT数据</param>
    /// <param name="startDate">开始日期</param>
    /// <param name="endDate">结束日期</param>
    /// <returns>生成的文件路径</returns>
    public async Task<string> GenerateExcelReportAsync(GMTReportResponse response, DateTime startDate, DateTime endDate)
    {
        // 检查模板文件是否存在
        if (!File.Exists(_templatePath))
        {
            throw new FileNotFoundException($"模板文件不存在: {_templatePath}");
        }

        // 生成输出文件名
        var fileName = $"GMT_Report_{startDate:yyyyMMdd}_{endDate:yyyyMMdd}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        var outputPath = Path.Combine(_outputDirectory, fileName);

        // 使用ClosedXML打开模板并填充数据
        using (var workbook = new XLWorkbook(_templatePath))
        {
            foreach (var deviceSummary in response.DeviceSummaries)
            {
                foreach (var cpData in deviceSummary.CPData)
                {
                    // 提取DEVICE前6位（如 NS3607_1B -> NS3607）
                    var baseDeviceName = deviceSummary.DeviceName.Length >= 6 
                        ? deviceSummary.DeviceName.Substring(0, 6) 
                        : deviceSummary.DeviceName;
                    
                    // 根据DEVICE前6位和CP号确定工作表名
                    var sheetName = $"{baseDeviceName}_{cpData.CPNumber}";
                    
                    Console.WriteLine($"[INFO] 查找工作表: {sheetName}");
                    
                    // 查找模板中的工作表
                    if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet))
                    {
                        Console.WriteLine($"[WARNING] 模板中未找到工作表 {sheetName}，跳过该设备");
                        continue;
                    }
                    
                    Console.WriteLine($"[INFO] 找到工作表 {sheetName}，开始填充数据");

                    // 填充数据（从第3行开始）
                    // 按Lot分组，每个Lot的Total行在前，数据行在后
                    int row = 3;
                    var lotGroups = cpData.DataRows.GroupBy(r => r.WF_Lot).OrderBy(g => g.Key);
                    
                    foreach (var lotGroup in lotGroups)
                    {
                        var lotRows = lotGroup.OrderBy(r => int.Parse(r.WF_No)).ToList();
                        var firstRow = lotRows.First();
                        
                        // 先写Total行（在数据前面）
                        int totalRow = row;
                        
                        // Total行：填充Lot信息（不合并）
                        worksheet.Cell(totalRow, 2).Value = "HTSH";
                        worksheet.Cell(totalRow, 3).Value = firstRow.WF_Lot;
                        worksheet.Cell(totalRow, 4).Value = firstRow.Tester;
                        worksheet.Cell(totalRow, 5).Value = firstRow.Temp;
                        worksheet.Cell(totalRow, 6).Value = firstRow.PRG_Name;
                        worksheet.Cell(totalRow, 7).Value = "Total";  // 第7列显示"Total"
                        
                        row++; // 移到下一行，开始填充数据
                        
                        int dataStartRow = row;
                        
                        // 填充该Lot的所有Wafer数据
                        foreach (var dataRow in lotRows)
                        {
                            // 日期格式转换为YYYY/MM/DD
                            var formattedDate = FormatDateToYYYYMMDD(dataRow.Date);
                            worksheet.Cell(row, 1).Value = formattedDate;           // 日期
                            // 列2-6先不合并，填充完数据后再统一合并
                            worksheet.Cell(row, 7).Value = dataRow.WF_No;          // Wafer ID
                            worksheet.Cell(row, 8).Value = dataRow.GrossQty;       // Gross Qty
                            
                            // 填充BIN数据（从第9列开始）
                            int binCol = 9;
                            foreach (var bin in dataRow.BinCounts.OrderBy(b => b.Key))
                            {
                                worksheet.Cell(row, binCol).Value = bin.Value;
                                binCol++;
                            }
                            
                            // Yield列
                            worksheet.Cell(row, binCol).Value = dataRow.Yield;
                            
                            row++;
                        }
                        
                        int dataEndRow = row - 1;
                        
                        // 合并当前Lot的所有数据行的列2-6（纵向合并）
                        if (dataStartRow <= dataEndRow)
                        {
                            worksheet.Range(dataStartRow, 2, dataEndRow, 6).Merge();
                        }
                        
                        // 回填Total行的公式（引用下面的数据行）
                        int totalBinCol = 8;
                        
                        // Gross Qty（第8列）使用SUM公式
                        worksheet.Cell(totalRow, totalBinCol).FormulaA1 = $"SUM({GetColumnLetter(totalBinCol)}{dataStartRow}:{GetColumnLetter(totalBinCol)}{dataEndRow})";
                        totalBinCol++;
                        
                        // 各个BIN列使用SUM公式
                        var binKeys = firstRow.BinCounts.Keys.OrderBy(k => k).ToList();
                        foreach (var binKey in binKeys)
                        {
                            worksheet.Cell(totalRow, totalBinCol).FormulaA1 = $"SUM({GetColumnLetter(totalBinCol)}{dataStartRow}:{GetColumnLetter(totalBinCol)}{dataEndRow})";
                            totalBinCol++;
                        }
                        
                        // Yield列使用AVERAGE公式
                        worksheet.Cell(totalRow, totalBinCol).FormulaA1 = $"AVERAGE({GetColumnLetter(totalBinCol)}{dataStartRow}:{GetColumnLetter(totalBinCol)}{dataEndRow})";
                        
                        // 设置Total行的格式：橙色背景
                        var totalRange = worksheet.Range(totalRow, 1, totalRow, totalBinCol);
                        totalRange.Style.Fill.BackgroundColor = XLColor.Orange;
                        totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        
                        // row已经指向下一个Lot的起始位置
                    }
                    
                    // 设置整个数据区域的格式：黄色背景和边框
                    if (cpData.DataRows.Any())
                    {
                        int lastDataCol = 9 + cpData.DataRows.First().BinCounts.Count;
                        int lastRow = row - 1; // 减1因为row已经移到了下一行
                        int firstRow = 3;
                        
                        // 为每个单元格单独设置格式（包括边框）
                        for (int r = firstRow; r <= lastRow; r++)
                        {
                            // 判断是否是Total行（第7列的值为"Total"）
                            bool isTotalRow = worksheet.Cell(r, 7).Value.ToString() == "Total";
                            
                            for (int c = 1; c <= lastDataCol; c++)
                            {
                                var cell = worksheet.Cell(r, c);
                                
                                // 设置边框
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                                
                                // 如果不是Total行，设置黄色背景
                                if (!isTotalRow)
                                {
                                    cell.Style.Fill.BackgroundColor = XLColor.Yellow;
                                }
                            }
                        }
                    }
                }
            }

            // 保存文件
            await Task.Run(() => workbook.SaveAs(outputPath));
        }

        Console.WriteLine($"[INFO] Excel文件已生成: {outputPath}");
        return outputPath;
    }

    /// <summary>
    /// 将列号转换为Excel列字母（如1->A, 27->AA）
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
