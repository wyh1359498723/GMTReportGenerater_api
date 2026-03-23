using System.ComponentModel.DataAnnotations;

namespace GMTReportGenerater_api.Models;

/// <summary>
/// GMT月报查询参数
/// </summary>
public class GMTReportRequest
{
    [Required]
    public string StartDate { get; set; } = string.Empty;

    [Required]
    public string EndDate { get; set; } = string.Empty;
}

/// <summary>
/// 新的GMT数据查询结果模型
/// </summary>
public class GMTDataRecord
{
    public string DEVICE { get; set; } = string.Empty;
    public string CP_NO { get; set; } = string.Empty;
    public string Lot_No { get; set; } = string.Empty;
    public string Tester_ID { get; set; } = string.Empty;
    public int Temp { get; set; }
    public string Test_Program { get; set; } = string.Empty;
    public string DateStr { get; set; } = string.Empty;
    public string Wafer_ID { get; set; } = string.Empty;
    public int Wafer_Total_Dies { get; set; }
    public string SBIN { get; set; } = string.Empty;
    public string Card_ID { get; set; } = string.Empty;
}

/// <summary>
/// Lot级别信息
/// </summary>
public class LotInfo
{
    public string Lot_No { get; set; } = string.Empty;
    public string Tester_ID { get; set; } = string.Empty;
    public int Temp { get; set; }
    public string Test_Program { get; set; } = string.Empty;
    public List<WaferInfo> Wafers { get; set; } = new();
}

/// <summary>
/// Wafer级别信息
/// </summary>
public class WaferInfo
{
    public string Date { get; set; } = string.Empty;
    public string Wafer_ID { get; set; } = string.Empty;
    public int Wafer_Total_Dies { get; set; }
    public Dictionary<string, int> BinCounts { get; set; } = new();
    public string Yield { get; set; } = string.Empty;
}

/// <summary>
/// Sheet配置信息
/// </summary>
public class SheetConfig
{
    public string SheetName { get; set; } = string.Empty; // 格式: {device}_{cp}
    public CPConfig CPConfig { get; set; } = new();
    public List<LotInfo> LotInfos { get; set; } = new();
}

/// <summary>
/// Device配置信息
/// </summary>
public class DeviceConfig
{
    public string DeviceName { get; set; } = string.Empty;
    public List<CPConfig> CPConfigs { get; set; } = new();
}

/// <summary>
/// CP配置信息
/// </summary>
public class CPConfig
{
    public string CPNumber { get; set; } = string.Empty;
    public int GrossQty { get; set; }
    public List<string> PassBins { get; set; } = new();
    public string CellBorder { get; set; } = string.Empty;
    public List<string> ShowBins { get; set; } = new();
}

/// <summary>
/// 处理后的数据行
/// </summary>
public class ProcessedDataRow
{
    public string Date { get; set; } = string.Empty;
    public string Tester { get; set; } = string.Empty;
    public string WF_Lot { get; set; } = string.Empty;
    public string Temp { get; set; } = string.Empty;
    public string PRG_Name { get; set; } = string.Empty;
    public string WF_No { get; set; } = string.Empty;
    public int GrossQty { get; set; }
    public Dictionary<string, int> BinCounts { get; set; } = new();
    public string Yield { get; set; } = string.Empty;
    public string Card_ID { get; set; } = string.Empty;
}

/// <summary>
/// Device数据汇总
/// </summary>
public class DeviceSummary
{
    public string DeviceName { get; set; } = string.Empty;
    public List<CPData> CPData { get; set; } = new();
}

/// <summary>
/// CP数据
/// </summary>
public class CPData
{
    public string CPNumber { get; set; } = string.Empty;
    public List<ProcessedDataRow> DataRows { get; set; } = new();
    public ProcessedDataRow? TotalRow { get; set; }
}

/// <summary>
/// GMT月报响应
/// </summary>
public class GMTReportResponse
{
    public List<DeviceSummary> DeviceSummaries { get; set; } = new();
    public string ReportDate { get; set; } = string.Empty;
    public string DateRange { get; set; } = string.Empty;
}

/// <summary>
/// 设备CP组合
/// </summary>
public class DeviceCPCombination
{
    public string Device { get; set; } = string.Empty;
    public string CP_No { get; set; } = string.Empty;
}

/// <summary>
/// GMT详细记录
/// </summary>
public class GMTDetailRecord
{
    public string Device { get; set; } = string.Empty;
    public string WF_Lot { get; set; } = string.Empty;
    public string Serial_No { get; set; } = string.Empty;
    public string CP_No { get; set; } = string.Empty;
    public string RP_No { get; set; } = string.Empty;
    public int Tested_WF_Qty { get; set; }
}

/// <summary>
/// GMT测试详情
/// </summary>
public class GMTTestDetail
{
    public string WF_No { get; set; } = string.Empty;
    public string Tester { get; set; } = string.Empty;
    public DateTime Start_Date { get; set; }
    public string SBIN { get; set; } = string.Empty;
    /// <summary>硬 Bin 串（QC 报表中部分机型用 HBIN 替代 SBIN）</summary>
    public string HBIN { get; set; } = string.Empty;
    public string Card_ID { get; set; } = string.Empty;
}
