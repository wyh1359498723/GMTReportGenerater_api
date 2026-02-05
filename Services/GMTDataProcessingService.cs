using GMTReportGenerater_api.Models;
using GMTReportGenerater_api.Services;

namespace GMTReportGenerater_api.Services;

/// <summary>
/// GMT数据处理服务 - 按照Delphi代码逻辑重新实现
/// </summary>
public class GMTDataProcessingService
{
    private readonly GMTConfigService _configService;
    private readonly GMTDatabaseService _databaseService;

    public GMTDataProcessingService(GMTConfigService configService, GMTDatabaseService databaseService)
    {
        _configService = configService;
        _databaseService = databaseService;
    }

    /// <summary>
    /// 处理GMT数据 - 使用新的数据结构和逻辑
    /// </summary>
    public async Task<GMTReportResponse> ProcessGMTDataAsync(DateTime startDate, DateTime endDate)
    {
        var response = new GMTReportResponse
        {
            ReportDate = DateTime.Now.ToString("yyyy-MM-dd"),
            DateRange = $"{startDate:yyyy-MM-dd} 至 {endDate:yyyy-MM-dd}"
        };

        try
        {
            // 查询所有GMT数据
            var rawData = await _databaseService.QueryGMTDataAsync(startDate, endDate);

            if (!rawData.Any())
            {
                return response;
            }

            // 按Device和CP分组处理，找到对应的Sheet
            var sheetGroups = rawData.GroupBy(r => new { r.DEVICE, r.CP_NO });

            foreach (var sheetGroup in sheetGroups)
            {
                var deviceName = sheetGroup.Key.DEVICE;
                var cpNumber = sheetGroup.Key.CP_NO;

                // 从设备名中提取基础名称（去掉版本后缀，如 NS3607_1B -> NS3607）
                var baseDeviceName = ExtractBaseDeviceName(deviceName);

                // 获取CP配置
                var cpConfig = _configService.GetCPConfig(baseDeviceName, cpNumber);
                if (cpConfig == null)
                {
                    Console.WriteLine($"未找到设备 {deviceName} (基础名:{baseDeviceName}) 的CP {cpNumber} 配置，跳过");
                    continue;
                }

                // 创建DeviceSummary
                var deviceSummary = response.DeviceSummaries.FirstOrDefault(d => d.DeviceName == deviceName);
                if (deviceSummary == null)
                {
                    deviceSummary = new DeviceSummary { DeviceName = deviceName };
                    response.DeviceSummaries.Add(deviceSummary);
                }

                var cpData = new CPData
                {
                    CPNumber = cpNumber
                };

                // 按Lot_No分组处理
                var lotGroups = sheetGroup.GroupBy(r => r.Lot_No);

                foreach (var lotGroup in lotGroups)
                {
                    var lotRecord = lotGroup.First(); // Lot级信息从第一条记录获取
                    var lotInfo = new LotInfo
                    {
                        Lot_No = lotRecord.Lot_No,
                        Tester_ID = lotRecord.Tester_ID,
                        Temp = lotRecord.Temp,
                        Test_Program = lotRecord.Test_Program,
                        Wafers = new List<WaferInfo>()
                    };

                    // 按Wafer_ID分组处理
                    var waferGroups = lotGroup.GroupBy(r => r.Wafer_ID);

                    foreach (var waferGroup in waferGroups)
                    {
                        var waferRecord = waferGroup.First(); // Wafer级信息从第一条记录获取

                        var waferInfo = new WaferInfo
                        {
                            Date = waferRecord.DateStr,
                            Wafer_ID = waferRecord.Wafer_ID,
                            Wafer_Total_Dies = waferRecord.Wafer_Total_Dies,
                            BinCounts = new Dictionary<string, int>(),
                            Yield = ""
                        };

                        var row = new ProcessedDataRow
                        {
                            Date = waferInfo.Date,
                            Tester = lotRecord.Tester_ID,
                            WF_Lot = lotRecord.Lot_No,
                            Temp = lotRecord.Temp.ToString(),
                            PRG_Name = lotRecord.Test_Program,
                            WF_No = waferInfo.Wafer_ID,
                            GrossQty = waferInfo.Wafer_Total_Dies,
                            BinCounts = new Dictionary<string, int>(),
                            Yield = "",
                            Card_ID = waferRecord.Card_ID
                        };

                        // 初始化所有显示bin的数量为0
                        foreach (var bin in cpConfig.ShowBins)
                        {
                            waferInfo.BinCounts[bin] = 0;
                            row.BinCounts[bin] = 0;
                        }

                        // 统计每个bin的数量
                        foreach (var record in waferGroup)
                        {
                            var binCounts = ParseSBINData(record.SBIN, cpConfig.ShowBins);
                            foreach (var binCount in binCounts)
                            {
                                if (waferInfo.BinCounts.ContainsKey(binCount.Key))
                                {
                                    waferInfo.BinCounts[binCount.Key] += binCount.Value;
                                    row.BinCounts[binCount.Key] += binCount.Value;
                                }
                            }
                        }

                        // 计算Yield
                        var passQty = 0;
                        var totalQty = 0;
                        foreach (var bin in cpConfig.ShowBins)
                        {
                            var count = row.BinCounts.GetValueOrDefault(bin, 0);
                            totalQty += count;
                            if (cpConfig.PassBins.Contains(bin))
                            {
                                passQty += count;
                            }
                        }

                        waferInfo.Yield = totalQty > 0 ? $"{(double)passQty / totalQty * 100:F1}%" : "0.0%";
                        row.Yield = waferInfo.Yield;

                        cpData.DataRows.Add(row);
                        lotInfo.Wafers.Add(waferInfo);
                    }
                }

                // 排序
                cpData.DataRows = cpData.DataRows
                    .OrderBy(r => r.WF_Lot)
                    .ThenBy(r => int.Parse(r.WF_No))
                    .ToList();

                if (cpData.DataRows.Any())
                {
                    // 计算汇总行
                    cpData.TotalRow = CalculateTotalRow(cpData.DataRows, cpConfig);
                    deviceSummary.CPData.Add(cpData);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理GMT数据时发生错误: {ex.Message}");
            throw;
        }

        return response;
    }

    /// <summary>
    /// 处理测试详情数据（参考Delphi代码第155-246行）
    /// </summary>
    private ProcessedDataRow? ProcessTestDetailsAsync(
        List<GMTTestDetail> testDetails,
        GMTDetailRecord detailRecord,
        string testProgram,
        CPConfig cpConfig)
    {
        try
        {
            // 使用StringBuilder来构建数据行（模拟Delphi的list.Add逻辑）
            var dataRows = new List<ProcessedDataRow>();

            foreach (var testDetail in testDetails)
            {
                if (string.IsNullOrEmpty(testDetail.SBIN))
                {
                    Console.WriteLine($"WF_LOT: {detailRecord.WF_Lot} 的第{testDetail.WF_No}片没有SOFT BIN值，请联系IT");
                    continue;
                }

                var dataRow = new ProcessedDataRow
                {
                    Date = testDetail.Start_Date.ToString("yyyy-MM-dd"),
                    Tester = testDetail.Tester,
                    WF_Lot = detailRecord.WF_Lot,
                    Temp = "25", // 固定温度
                    PRG_Name = testProgram,
                    WF_No = testDetail.WF_No,
                    GrossQty = cpConfig.GrossQty,
                    BinCounts = new Dictionary<string, int>(),
                    Yield = ""
                };

                // 解析SBIN数据（参考Delphi代码第182-214行）
                var binCounts = ParseSBINData(testDetail.SBIN, cpConfig.ShowBins);

                // 初始化所有显示bin的数量为0
                foreach (var bin in cpConfig.ShowBins)
                {
                    dataRow.BinCounts[bin] = 0;
                }

                // 填充实际数量
                foreach (var binCount in binCounts)
                {
                    if (dataRow.BinCounts.ContainsKey(binCount.Key))
                    {
                        dataRow.BinCounts[binCount.Key] = binCount.Value;
                    }
                }

                // 计算yield（参考Delphi代码第217-219行）
                var passQty = 0;
                var totalQty = 0;

                foreach (var bin in cpConfig.ShowBins)
                {
                    var count = dataRow.BinCounts.GetValueOrDefault(bin, 0);
                    totalQty += count;

                    // 检查是否为合格bin（参考Delphi代码第202行）
                    // 合格bin判断：pos('$'+gmt_tmp1+'$',gmt_pass_bin)>0
                    if (cpConfig.PassBins.Contains(bin))
                    {
                        passQty += count;
                    }
                }

                dataRow.Yield = totalQty > 0 ? $"{(double)passQty / totalQty * 100:F1}%" : "0.0%";

                // 添加gross qty到数据行（参考Delphi代码第222-223行）
                // 在Delphi代码中，gross qty被添加到数据行的末尾
                dataRow.GrossQty = totalQty; // 使用实际计算出的总数

                dataRows.Add(dataRow);
            }

                // 如果有多个WF，按WF_LOT和PRG_NAME分组合并（参考Delphi代码第228-237行）
                if (dataRows.Count > 1)
                {
                    var mergedRows = MergeDataRows(dataRows, cpConfig);
                    return mergedRows.FirstOrDefault();
                }

            return dataRows.FirstOrDefault();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理测试详情时发生错误: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 解析SBIN数据（参考Delphi代码第182-214行）
    /// SBIN格式: |bin1:count1|bin2:count2|...
    /// </summary>
    private Dictionary<string, int> ParseSBINData(string sbin, List<string> showBins)
    {
        var binCounts = new Dictionary<string, int>();

        try
        {
            // SBIN格式示例: |2:150|4:200|5:50|...
            var parts = sbin.Split('|').Where(p => !string.IsNullOrEmpty(p)).ToArray();

            foreach (var part in parts)
            {
                var colonPos = part.IndexOf(':');
                if (colonPos > 0)
                {
                    var binStr = part.Substring(0, colonPos);
                    var countStr = part.Substring(colonPos + 1);

                    if (int.TryParse(binStr, out var bin) && int.TryParse(countStr, out var count))
                    {
                        binCounts[binStr] = count;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"解析SBIN数据失败: {sbin}, 错误: {ex.Message}");
        }

        return binCounts;
    }

    /// <summary>
    /// 合并数据行（参考Delphi代码第228-237行）
    /// </summary>
    private List<ProcessedDataRow> MergeDataRows(List<ProcessedDataRow> rows, CPConfig cpConfig)
    {
        // 按WF_LOT和PRG_NAME分组
        var groupedRows = rows.GroupBy(r => $"{r.WF_Lot}_{r.PRG_Name}_{r.WF_No}");

        var mergedRows = new List<ProcessedDataRow>();

        foreach (var group in groupedRows)
        {
            var firstRow = group.First();
            var mergedRow = new ProcessedDataRow
            {
                Date = firstRow.Date,
                Tester = firstRow.Tester,
                WF_Lot = firstRow.WF_Lot,
                Temp = firstRow.Temp,
                PRG_Name = firstRow.PRG_Name,
                WF_No = firstRow.WF_No,
                GrossQty = group.Sum(r => r.GrossQty),
                BinCounts = new Dictionary<string, int>(),
                Yield = ""
            };

            // 合并bin数量
            foreach (var bin in firstRow.BinCounts.Keys)
            {
                mergedRow.BinCounts[bin] = group.Sum(r => r.BinCounts.GetValueOrDefault(bin, 0));
            }

            // 重新计算yield
            var passQty = 0;
            var totalQty = 0;
            foreach (var bin in mergedRow.BinCounts)
            {
                totalQty += bin.Value;
                // 检查是否为合格bin（使用CP配置）
                if (cpConfig.PassBins.Contains(bin.Key))
                {
                    passQty += bin.Value;
                }
            }

            mergedRow.Yield = totalQty > 0 ? $"{(double)passQty / totalQty * 100:F1}%" : "0.0%";

            mergedRows.Add(mergedRow);
        }

        return mergedRows;
    }

    /// <summary>
    /// 计算汇总行（参考Delphi代码第275-302行）
    /// </summary>
    private ProcessedDataRow CalculateTotalRow(List<ProcessedDataRow> rows, CPConfig cpConfig)
    {
        if (!rows.Any())
            return new ProcessedDataRow();

        var totalRow = new ProcessedDataRow
        {
            Date = "",
            Tester = "",
            WF_Lot = "",
            Temp = "",
            PRG_Name = "",
            WF_No = "",
            GrossQty = cpConfig.GrossQty,
            BinCounts = new Dictionary<string, int>(),
            Yield = ""
        };

        // 初始化bin数量
        foreach (var bin in rows.First().BinCounts.Keys)
        {
            totalRow.BinCounts[bin] = 0;
        }

        // 汇总所有行的bin数量
        foreach (var row in rows)
        {
            foreach (var binCount in row.BinCounts)
            {
                if (totalRow.BinCounts.ContainsKey(binCount.Key))
                {
                    totalRow.BinCounts[binCount.Key] += binCount.Value;
                }
            }
        }

        // 计算总yield
        var totalPassQty = 0;
        var totalQty = 0;

        foreach (var bin in totalRow.BinCounts)
        {
            totalQty += bin.Value;
            if (cpConfig.PassBins.Contains(bin.Key))
            {
                totalPassQty += bin.Value;
            }
        }

        totalRow.Yield = totalQty > 0 ? $"{(double)totalPassQty / totalQty * 100:F1}%" : "0.0%";

        return totalRow;
    }

    /// <summary>
    /// 从设备名中提取基础名称（去掉版本后缀）
    /// 例如: NS3607_1B -> NS3607, NS3608_1C -> NS3608
    /// </summary>
    private string ExtractBaseDeviceName(string deviceName)
    {
        if (string.IsNullOrEmpty(deviceName))
        {
            return deviceName;
        }

        // 查找下划线位置
        var underscoreIndex = deviceName.IndexOf('_');
        if (underscoreIndex > 0)
        {
            // 返回下划线之前的部分
            return deviceName.Substring(0, underscoreIndex);
        }

        // 如果没有下划线，返回原始名称
        return deviceName;
    }
}
