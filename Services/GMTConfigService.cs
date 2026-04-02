using GMTReportGenerater_api.Models;

namespace GMTReportGenerater_api.Services;

/// <summary>
/// GMT配置服务
/// </summary>
public class GMTConfigService
{
    private readonly string _configPath;
    private Dictionary<string, DeviceConfig> _deviceConfigs = new();
    private List<string> _gmtDevices = new();
    private List<string> _gmtDevicesAll = new();
    /// <summary>[GMT] DEVICE_ALL 原始串，用于 pos(device, gmt_device_all) 子串判断</summary>
    private string _gmtDeviceAllRaw = string.Empty;
    /// <summary>DEVICE_HBIN 原始串（Delphi：pos(设备前缀, gmt_hbin)&gt;0 时用 HBIN）</summary>
    private string _gmtDeviceHbinRaw = string.Empty;
    private string _qcTemplateFileName = string.Empty;

    public GMTConfigService(IWebHostEnvironment environment)
    {
        _configPath = Path.Combine(environment.ContentRootPath, "Template", "setup.ini");
        LoadConfiguration();
    }

    /// <summary>
    /// 加载配置文件
    /// </summary>
    private void LoadConfiguration()
    {
        if (!File.Exists(_configPath))
        {
            throw new FileNotFoundException($"配置文件不存在: {_configPath}");
        }

        var lines = File.ReadAllLines(_configPath);
        var currentSection = string.Empty;

        foreach (var line in lines)
        {
            var trimmedLine = line.Trim();
            if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";"))
                continue;

            if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
            {
                currentSection = trimmedLine.Trim('[', ']');
                continue;
            }

            if (currentSection == "GMT" && trimmedLine.Contains("="))
            {
                var parts = trimmedLine.Split('=', 2);
                var key = parts[0].Trim();
                var value = parts[1].Trim();

                if (key == "DEVICE")
                {
                    _gmtDevices = value.Split(',').Select(d => d.Trim()).ToList();
                }
                else if (key == "DEVICE_ALL")
                {
                    _gmtDeviceAllRaw = value.Trim();
                    _gmtDevicesAll = value.Split(',').Select(d => d.Trim()).Where(d => d.Length > 0).ToList();
                }
                else if (key == "DEVICE_HBIN")
                {
                    _gmtDeviceHbinRaw = value.Trim();
                }
            }
            else if (currentSection == "BASIC" && trimmedLine.Contains("="))
            {
                // 仅解析 QC 模板键，其余 BASIC 行忽略
                var partsBasic = trimmedLine.Split('=', 2);
                var k = partsBasic[0].Trim().Trim('\'', '"');
                if (k.Equals("GMT-BB30-SPECIAL2", StringComparison.OrdinalIgnoreCase) && partsBasic.Length > 1)
                {
                    _qcTemplateFileName = partsBasic[1].Trim().Trim('\'', '"');
                }
            }
            else if (!string.IsNullOrEmpty(currentSection) && trimmedLine.Contains("="))
            {
                var parts = trimmedLine.Split('=', 2);
                var cpKey = parts[0].Trim();
                var cpValue = parts[1].Trim();

                if (!cpKey.StartsWith("CP"))
                    continue;

                if (!_deviceConfigs.ContainsKey(currentSection))
                {
                    _deviceConfigs[currentSection] = new DeviceConfig
                    {
                        DeviceName = currentSection
                    };
                }

                var cpConfig = ParseCPConfig(cpKey, cpValue);
                _deviceConfigs[currentSection].CPConfigs.Add(cpConfig);
            }
        }
    }

    /// <summary>
    /// 解析CP配置 - 按照Delphi代码逻辑
    /// 格式: 总数*$合格bin$&单元格边界|显示bin列表
    /// 例如: 19873*$2$&AX|2,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,
    /// </summary>
    private CPConfig ParseCPConfig(string cpKey, string cpValue)
    {
        var config = new CPConfig
        {
            CPNumber = cpKey
        };

        try
        {
            // 解析总数 (gmt_gross_qty)
            var asteriskPos = cpValue.IndexOf('*');
            if (asteriskPos > 0)
            {
                var grossQtyStr = cpValue.Substring(0, asteriskPos);
                if (int.TryParse(grossQtyStr, out var grossQty))
                {
                    config.GrossQty = grossQty;
                }
            }

            // 解析合格bin (gmt_pass_bin)
            var dollarPos = cpValue.IndexOf('$', asteriskPos + 1);
            var ampPos = cpValue.IndexOf('&', dollarPos + 1);
            if (dollarPos > 0 && ampPos > dollarPos)
            {
                var passBinStr = cpValue.Substring(dollarPos + 1, ampPos - dollarPos - 1);
                if (!string.IsNullOrEmpty(passBinStr))
                {
                    config.PassBins = passBinStr.Split('$').Where(b => !string.IsNullOrEmpty(b)).ToList();
                }
            }

            // 解析单元格边界 (gmt_cell_border)
            var pipePos = cpValue.IndexOf('|', ampPos + 1);
            if (ampPos > 0 && pipePos > ampPos)
            {
                config.CellBorder = cpValue.Substring(ampPos + 1, pipePos - ampPos - 1);
            }

            // 解析显示bin列表 (gmt_show_bin)
            if (pipePos > 0)
            {
                var showBinsStr = cpValue.Substring(pipePos + 1).TrimEnd(',');
                config.ShowBins = showBinsStr.Split(',').Where(b => !string.IsNullOrEmpty(b)).ToList();
            }
        }
        catch (Exception ex)
        {
            // 解析失败时记录日志但不抛出异常
            Console.WriteLine($"解析CP配置失败 {cpKey}: {cpValue}, 错误: {ex.Message}");
        }

        return config;
    }

    /// <summary>
    /// 获取GMT设备列表
    /// </summary>
    public List<string> GetGMTDevices()
    {
        return _gmtDevices;
    }

    /// <summary>
    /// 获取GMT设备全部列表（包括特殊设备）
    /// </summary>
    public List<string> GetGMTDevicesAll()
    {
        return _gmtDevicesAll;
    }

    /// <summary>DEVICE_ALL 原始配置值（Delphi pos(trim(device), gmt_device_all)）</summary>
    public string GetGMTDeviceAllRaw() => _gmtDeviceAllRaw;

    /// <summary>
    /// 获取设备配置
    /// </summary>
    public DeviceConfig? GetDeviceConfig(string deviceName)
    {
        return _deviceConfigs.ContainsKey(deviceName) ? _deviceConfigs[deviceName] : null;
    }

    /// <summary>
    /// 获取CP配置
    /// </summary>
    public CPConfig? GetCPConfig(string deviceName, string cpNumber)
    {
        var deviceConfig = GetDeviceConfig(deviceName);
        return deviceConfig?.CPConfigs.FirstOrDefault(cp => cp.CPNumber == cpNumber);
    }

    /// <summary>
    /// 获取所有设备配置
    /// </summary>
    public Dictionary<string, DeviceConfig> GetAllDeviceConfigs()
    {
        return _deviceConfigs;
    }

    /// <summary>
    /// Delphi pos(gmt_device, gmt_hbin)&gt;0：当前 6 位 device 前缀是否出现在 DEVICE_HBIN 配置串中。
    /// </summary>
    public bool UseHardBinForDevicePrefix(string deviceSixCharPrefix)
    {
        if (string.IsNullOrEmpty(deviceSixCharPrefix) || string.IsNullOrEmpty(_gmtDeviceHbinRaw))
            return false;
        return _gmtDeviceHbinRaw.IndexOf(deviceSixCharPrefix, StringComparison.Ordinal) >= 0;
    }

    /// <summary>setup.ini [BASIC] 中 GMT-BB30-SPECIAL2 对应的模板文件名</summary>
    public string GetQcTemplateFileName() => _qcTemplateFileName;

    /// <summary>QC 模板完整路径（ContentRoot/Template/文件名）</summary>
    public string GetQcTemplateFullPath(string contentRootPath)
    {
        if (string.IsNullOrWhiteSpace(_qcTemplateFileName))
            throw new InvalidOperationException("setup.ini [BASIC] 未配置 GMT-BB30-SPECIAL2 模板文件名");
        return Path.Combine(contentRootPath, "Template", _qcTemplateFileName);
    }

    /// <summary>QC 模板第 1 行中 Yield 列的表头文字；仅当列名与此匹配（忽略大小写）时才写入 Yield，并与 Bin 列起始位置对齐。</summary>
    public string GetQcTotalYieldColumnHeader() => "Total yield";
}
