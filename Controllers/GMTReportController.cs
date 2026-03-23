using Microsoft.AspNetCore.Mvc;
using GMTReportGenerater_api.Models;
using GMTReportGenerater_api.Services;
using System.Globalization;

namespace GMTReportGenerater_api.Controllers;

[ApiController]
[Route("api/[controller]")]
public class GMTReportController : ControllerBase
{
    private readonly GMTConfigService _configService;
    private readonly GMTDataProcessingService _dataProcessingService;
    private readonly GMTDatabaseService _databaseService;
    private readonly GMTExcelGenerationService _excelGenerationService;
    private readonly GMTExcelGenerationSpecialService _excelGenerationSpecialService;
    private readonly GMTExcelGenerationQcService _excelGenerationQcService;

    public GMTReportController(
        GMTConfigService configService,
        GMTDataProcessingService dataProcessingService,
        GMTDatabaseService databaseService,
        GMTExcelGenerationService excelGenerationService,
        GMTExcelGenerationSpecialService excelGenerationSpecialService,
        GMTExcelGenerationQcService excelGenerationQcService)
    {
        _configService = configService;
        _dataProcessingService = dataProcessingService;
        _databaseService = databaseService;
        _excelGenerationService = excelGenerationService;
        _excelGenerationSpecialService = excelGenerationSpecialService;
        _excelGenerationQcService = excelGenerationQcService;
    }

    /// <summary>
    /// 智能日期解析 - 支持多种格式
    /// </summary>
    private bool TryParseFlexibleDate(string dateStr, out DateTime result)
    {
        // 支持的日期格式列表
        string[] formats = new[]
        {
            "yyyy-MM-dd HH:mm:ss",           // 标准格式
            "yyyy-MM-ddTHH:mm:ss.fffZ",      // ISO 8601 with milliseconds
            "yyyy-MM-ddTHH:mm:ssZ",          // ISO 8601 without milliseconds
            "yyyy-MM-ddTHH:mm:ss.fff",       // ISO 8601 local with milliseconds
            "yyyy-MM-ddTHH:mm:ss",           // ISO 8601 local
            "yyyy-MM-dd",                    // 仅日期
        };

        // 尝试使用指定格式解析
        if (DateTime.TryParseExact(
                dateStr,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                out result))
        {
            return true;
        }

        // 尝试通用解析（自动识别格式）
        if (DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out result))
        {
            return true;
        }

        result = DateTime.MinValue;
        return false;
    }

    /// <summary>
    /// 生成GMT月报（返回JSON数据）
    /// </summary>
    /// <param name="request">报告请求参数</param>
    /// <returns>GMT月报数据</returns>
    [HttpPost("generate")]
    public async Task<ActionResult<GMTReportResponse>> GenerateGMTReport([FromBody] GMTReportRequest request)
    {
        try
        {
            // 使用智能日期解析
            if (!TryParseFlexibleDate(request.StartDate, out DateTime startDate))
            {
                return BadRequest($"开始时间格式无效: {request.StartDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-01T16:00:00.000Z）");
            }

            if (!TryParseFlexibleDate(request.EndDate, out DateTime endDate))
            {
                return BadRequest($"结束时间格式无效: {request.EndDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-08T15:59:59.999Z）");
            }

            // 验证输入参数
            if (startDate >= endDate)
            {
                return BadRequest("开始时间必须小于结束时间");
            }

            if ((endDate - startDate).TotalDays > 365)
            {
                return BadRequest("时间范围不能超过一年");
            }

            // 使用新的三步查询逻辑处理数据
            var response = await _dataProcessingService.ProcessGMTDataAsync(startDate, endDate);

            if (!response.DeviceSummaries.Any())
            {
                return NotFound("指定时间范围内没有找到GMT数据");
            }

            return Ok(response);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成报告时发生错误: {ex.Message}"+ Environment.NewLine+$"{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 生成并下载GMT月报Excel文件（POST方法，从Body读取参数）
    /// </summary>
    /// <param name="request">报告请求参数</param>
    /// <returns>Excel文件</returns>
    [HttpPost("generate-excel")]
    public async Task<IActionResult> GenerateGMTExcelReport([FromBody] GMTReportRequest request)
    {
        try
        {
            // 使用智能日期解析
            if (!TryParseFlexibleDate(request.StartDate, out DateTime startDate))
            {
                return BadRequest($"开始时间格式无效: {request.StartDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-01T16:00:00.000Z）");
            }

            if (!TryParseFlexibleDate(request.EndDate, out DateTime endDate))
            {
                return BadRequest($"结束时间格式无效: {request.EndDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-08T15:59:59.999Z）");
            }

            // 验证输入参数
            if (startDate >= endDate)
            {
                return BadRequest("开始时间必须小于结束时间");
            }

            if ((endDate - startDate).TotalDays > 365)
            {
                return BadRequest("时间范围不能超过一年");
            }

            // 处理数据
            var response = await _dataProcessingService.ProcessGMTDataAsync(startDate, endDate);

            if (!response.DeviceSummaries.Any())
            {
                return NotFound("指定时间范围内没有找到GMT数据");
            }

            // 生成Excel文件
            var filePath = await _excelGenerationService.GenerateExcelReportAsync(response, startDate, endDate);

            // 读取文件内容
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);

            // 返回文件
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成报告时发生错误: {ex.Message}" + Environment.NewLine + $"{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 生成并下载GMT月报Excel文件（GET方法，从URL查询参数读取）
    /// </summary>
    /// <param name="startDate">开始时间（支持 yyyy-MM-dd HH:mm:ss 或 ISO 8601 格式，如 2026-02-01T16:00:00.000Z）</param>
    /// <param name="endDate">结束时间（支持 yyyy-MM-dd HH:mm:ss 或 ISO 8601 格式，如 2026-02-08T15:59:59.999Z）</param>
    /// <returns>Excel文件</returns>
    [HttpGet("generate-excel")]
    public async Task<IActionResult> GenerateGMTExcelReportFromQuery([FromQuery] string startDate, [FromQuery] string endDate)
    {
        try
        {
            Console.WriteLine($"[INFO] GET请求参数 - StartDate: {startDate}, EndDate: {endDate}");

            // 使用智能日期解析
            if (!TryParseFlexibleDate(startDate, out DateTime parsedStartDate))
            {
                return BadRequest($"开始时间格式无效: {startDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-01T16:00:00.000Z）");
            }

            if (!TryParseFlexibleDate(endDate, out DateTime parsedEndDate))
            {
                return BadRequest($"结束时间格式无效: {endDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-08T15:59:59.999Z）");
            }

            Console.WriteLine($"[INFO] 解析后的日期 - StartDate: {parsedStartDate:yyyy-MM-dd HH:mm:ss}, EndDate: {parsedEndDate:yyyy-MM-dd HH:mm:ss}");

            // 验证输入参数
            if (parsedStartDate >= parsedEndDate)
            {
                return BadRequest("开始时间必须小于结束时间");
            }

            if ((parsedEndDate - parsedStartDate).TotalDays > 365)
            {
                return BadRequest("时间范围不能超过一年");
            }

            // 处理数据
            var response = await _dataProcessingService.ProcessGMTDataAsync(parsedStartDate, parsedEndDate);

            if (!response.DeviceSummaries.Any())
            {
                return NotFound("指定时间范围内没有找到GMT数据");
            }

            // 生成Excel文件
            var filePath = await _excelGenerationService.GenerateExcelReportAsync(response, parsedStartDate, parsedEndDate);

            // 读取文件内容
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);

            // 返回文件
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成报告时发生错误: {ex.Message}" + Environment.NewLine + $"{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 生成并下载GMT月报Excel文件 - Special格式（POST方法）
    /// </summary>
    /// <param name="request">报告请求参数</param>
    /// <returns>Excel文件</returns>
    [HttpPost("generate-excel-special")]
    public async Task<IActionResult> GenerateGMTExcelReportSpecial([FromBody] GMTReportRequest request)
    {
        try
        {
            // 使用智能日期解析
            if (!TryParseFlexibleDate(request.StartDate, out DateTime startDate))
            {
                return BadRequest($"开始时间格式无效: {request.StartDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-01T16:00:00.000Z）");
            }

            if (!TryParseFlexibleDate(request.EndDate, out DateTime endDate))
            {
                return BadRequest($"结束时间格式无效: {request.EndDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-08T15:59:59.999Z）");
            }

            // 验证输入参数
            if (startDate >= endDate)
            {
                return BadRequest("开始时间必须小于结束时间");
            }

            if ((endDate - startDate).TotalDays > 365)
            {
                return BadRequest("时间范围不能超过一年");
            }

            // 处理数据（使用相同的数据处理逻辑）
            var response = await _dataProcessingService.ProcessGMTDataAsync(startDate, endDate);

            if (!response.DeviceSummaries.Any())
            {
                return NotFound("指定时间范围内没有找到GMT数据");
            }

            // 生成Special格式Excel文件
            var filePath = await _excelGenerationSpecialService.GenerateExcelReportAsync(response, startDate, endDate);

            // 读取文件内容
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);

            // 返回文件
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成Special格式报告时发生错误: {ex.Message}" + Environment.NewLine + $"{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 生成并下载GMT月报Excel文件 - Special格式（GET方法，从URL查询参数读取）
    /// </summary>
    /// <param name="startDate">开始时间（支持 yyyy-MM-dd HH:mm:ss 或 ISO 8601 格式，如 2026-02-01T16:00:00.000Z）</param>
    /// <param name="endDate">结束时间（支持 yyyy-MM-dd HH:mm:ss 或 ISO 8601 格式，如 2026-02-08T15:59:59.999Z）</param>
    /// <returns>Excel文件</returns>
    [HttpGet("generate-excel-special")]
    public async Task<IActionResult> GenerateGMTExcelReportSpecialFromQuery([FromQuery] string startDate, [FromQuery] string endDate)
    {
        try
        {
            Console.WriteLine($"[INFO] GET请求参数（Special） - StartDate: {startDate}, EndDate: {endDate}");

            // 使用智能日期解析
            if (!TryParseFlexibleDate(startDate, out DateTime parsedStartDate))
            {
                return BadRequest($"开始时间格式无效: {startDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-01T16:00:00.000Z）");
            }

            if (!TryParseFlexibleDate(endDate, out DateTime parsedEndDate))
            {
                return BadRequest($"结束时间格式无效: {endDate}。支持格式：yyyy-MM-dd HH:mm:ss 或 ISO 8601（如 2026-02-08T15:59:59.999Z）");
            }

            Console.WriteLine($"[INFO] 解析后的日期（Special） - StartDate: {parsedStartDate:yyyy-MM-dd HH:mm:ss}, EndDate: {parsedEndDate:yyyy-MM-dd HH:mm:ss}");

            // 验证输入参数
            if (parsedStartDate >= parsedEndDate)
            {
                return BadRequest("开始时间必须小于结束时间");
            }

            if ((parsedEndDate - parsedStartDate).TotalDays > 365)
            {
                return BadRequest("时间范围不能超过一年");
            }

            // 处理数据（使用相同的数据处理逻辑）
            var response = await _dataProcessingService.ProcessGMTDataAsync(parsedStartDate, parsedEndDate);

            if (!response.DeviceSummaries.Any())
            {
                return NotFound("指定时间范围内没有找到GMT数据");
            }

            // 生成Special格式Excel文件
            var filePath = await _excelGenerationSpecialService.GenerateExcelReportAsync(response, parsedStartDate, parsedEndDate);

            // 读取文件内容
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);

            // 返回文件
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成Special格式报告时发生错误: {ex.Message}" + Environment.NewLine + $"{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 生成并下载 GMT QC 报表 Excel（POST，与 Delphi suiButton3Click / SPECIAL2 模板逻辑一致）
    /// </summary>
    [HttpPost("generate-excel-qc")]
    public async Task<IActionResult> GenerateGMTExcelQc([FromBody] GMTReportRequest request)
    {
        try
        {
            if (!TryParseFlexibleDate(request.StartDate, out DateTime startDate))
                return BadRequest($"开始时间格式无效: {request.StartDate}。");
            if (!TryParseFlexibleDate(request.EndDate, out DateTime endDate))
                return BadRequest($"结束时间格式无效: {request.EndDate}。");
            if (startDate >= endDate)
                return BadRequest("开始时间必须小于结束时间");
            if ((endDate - startDate).TotalDays > 365)
                return BadRequest("时间范围不能超过一年");

            var filePath = await _excelGenerationQcService.GenerateQcReportAsync(startDate, endDate);
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (InvalidOperationException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return NotFound(ex.Message);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成 QC 报表时发生错误: {ex.Message}" + Environment.NewLine + ex.StackTrace);
        }
    }

    /// <summary>
    /// 生成并下载 GMT QC 报表 Excel（GET 查询参数）
    /// </summary>
    [HttpGet("generate-excel-qc")]
    public async Task<IActionResult> GenerateGMTExcelQcFromQuery([FromQuery] string startDate, [FromQuery] string endDate)
    {
        try
        {
            if (!TryParseFlexibleDate(startDate, out DateTime parsedStart))
                return BadRequest($"开始时间格式无效: {startDate}。");
            if (!TryParseFlexibleDate(endDate, out DateTime parsedEnd))
                return BadRequest($"结束时间格式无效: {endDate}。");
            if (parsedStart >= parsedEnd)
                return BadRequest("开始时间必须小于结束时间");
            if ((parsedEnd - parsedStart).TotalDays > 365)
                return BadRequest("时间范围不能超过一年");

            var filePath = await _excelGenerationQcService.GenerateQcReportAsync(parsedStart, parsedEnd);
            var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            var fileName = Path.GetFileName(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (InvalidOperationException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return NotFound(ex.Message);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"生成 QC 报表时发生错误: {ex.Message}" + Environment.NewLine + ex.StackTrace);
        }
    }

    /// <summary>
    /// 获取GMT设备配置信息
    /// </summary>
    /// <returns>设备配置信息</returns>
    [HttpGet("devices")]
    public ActionResult<object> GetGMTDevices()
    {
        try
        {
            var devices = _configService.GetGMTDevices();
            var devicesAll = _configService.GetGMTDevicesAll();
            var allConfigs = _configService.GetAllDeviceConfigs();

            return Ok(new
            {
                GMTDevices = devices,
                GMTDevicesAll = devicesAll,
                DeviceConfigs = allConfigs
            });
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"获取设备配置时发生错误: {ex.Message}" + Environment.NewLine + $"{ex.StackTrace}");
        }
    }

}
