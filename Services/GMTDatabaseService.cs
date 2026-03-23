using System.Data;
using System.Globalization;
using GMTReportGenerater_api.Models;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;

namespace GMTReportGenerater_api.Services;

/// <summary>
/// GMT数据库服务
/// </summary>
public class GMTDatabaseService
{
    private readonly string _connectionString;

    public GMTDatabaseService(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("OracleDb")
            ?? throw new ArgumentNullException("OracleDb connection string not found");
    }

    private static string ConvertToDateString(object value)
    {
        if (value == null || value is DBNull)
        {
            return string.Empty;
        }

        if (value is DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd");
        }

        var str = Convert.ToString(value);
        return string.IsNullOrWhiteSpace(str) ? string.Empty : str.Trim();
    }

    private static int ConvertToInt32(object value)
    {
        if (value == null || value is DBNull)
        {
            return 0;
        }

        if (value is int intValue)
        {
            return intValue;
        }

        if (value is decimal decimalValue)
        {
            return (int)decimalValue;
        }

        if (value is long longValue)
        {
            return (int)longValue;
        }

        var str = Convert.ToString(value);
        return int.TryParse(str, out var parsed) ? parsed : 0;
    }

    private static string ConvertToString(object value)
    {
        if (value == null || value is DBNull)
        {
            return string.Empty;
        }

        return Convert.ToString(value)?.Trim() ?? string.Empty;
    }

    /// <summary>
    /// Oracle 中 start_date 可能为 DATE、TIMESTAMP、VARCHAR2 等，不能对非日期类型调用 GetDateTime。
    /// </summary>
    private static DateTime ConvertOracleValueToDateTime(object? value)
    {
        if (value == null || value is DBNull)
            return DateTime.MinValue;

        if (value is DateTime dt)
            return dt;
        if (value is DateTimeOffset dto)
            return dto.DateTime;

        switch (value)
        {
            case OracleDate od:
                return od.IsNull ? DateTime.MinValue : od.Value;
            case OracleTimeStamp ts:
                return ts.IsNull ? DateTime.MinValue : ts.Value;
            case OracleTimeStampTZ tstz:
                return tstz.IsNull ? DateTime.MinValue : tstz.Value;
            case OracleTimeStampLTZ ltz:
                return ltz.IsNull ? DateTime.MinValue : ltz.Value;
        }

        var s = Convert.ToString(value)?.Trim() ?? string.Empty;
        if (s.Length == 0)
            return DateTime.MinValue;

        string[] formats =
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss.ff",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "dd-MMM-yy",
            "dd-MMM-yyyy",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.fff"
        };

        if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exact))
            return exact;
        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
            return parsed;
        if (s.Length >= 10 && DateTime.TryParseExact(s.AsSpan(0, 10), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out exact))
            return exact;

        return DateTime.MinValue;
    }

    /// <summary>
    /// 查询GMT数据 - 使用新的SQL结构
    /// </summary>
    public async Task<List<GMTDataRecord>> QueryGMTDataAsync(DateTime startDate, DateTime endDate)
    {
        var data = new List<GMTDataRecord>();

        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = @"
                SELECT
                    h.DEVICE,
                    h.CP_NO,
                    h.WF_LOT as Lot_No,
                    p.TESTER as Tester_ID,
                    h.WF_QTY as Temp,
                    g.TEST_P as Test_Program,
                    h.OPEN_DATE as OpenDate,
                    p.WF_NO as Wafer_ID,
                    p.GROSS_QTY as Wafer_Total_Dies,
                    p.SBIN,
                    p.CARD_ID
                FROM p_data_detail p
                LEFT JOIN rtm_admin.rtm_p_data_head h
                    ON p.SERIAL_NO = h.SERIAL_NO
                LEFT JOIN RTM_ADMIN.RTM_P_DATA_SPLIT g
                    ON p.SERIAL_NO = g.SERIAL_NO
                WHERE
                    p.lot_id LIKE 'GMT%'
                    AND h.STATUS = 'FINISH'
                    AND SUBSTR(p.end_date, 1, 10) >= SUBSTR(:startDate, 1, 10)
                    AND SUBSTR(p.end_date, 1, 10) <= SUBSTR(:endDate, 1, 10)
                    AND h.HT_CODE = 'BB30'
                ORDER BY
                    p.LOT_ID,
                    TO_NUMBER(p.WF_NO)";

            using (var command = new OracleCommand(sql, connection))
            {
                var startDateStr = startDate.ToString("yyyy-MM-dd HH:mm:ss");
                var endDateStr = endDate.ToString("yyyy-MM-dd HH:mm:ss");
                
                command.Parameters.Add(":startDate", OracleDbType.Varchar2).Value = startDateStr;
                command.Parameters.Add(":endDate", OracleDbType.Varchar2).Value = endDateStr;

                Console.WriteLine($"[DEBUG] 查询参数: startDate={startDateStr}, endDate={endDateStr}");
                Console.WriteLine($"[DEBUG] SQL: {sql}");

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        var record = new GMTDataRecord
                        {
                            DEVICE = ConvertToString(reader.GetValue(reader.GetOrdinal("DEVICE"))),
                            CP_NO = ConvertToString(reader.GetValue(reader.GetOrdinal("CP_NO"))),
                            Lot_No = ConvertToString(reader.GetValue(reader.GetOrdinal("Lot_No"))),
                            Tester_ID = ConvertToString(reader.GetValue(reader.GetOrdinal("Tester_ID"))),
                            Temp = ConvertToInt32(reader.GetValue(reader.GetOrdinal("Temp"))),
                            Test_Program = reader.IsDBNull(reader.GetOrdinal("Test_Program"))
                                ? string.Empty
                                : reader.GetString(reader.GetOrdinal("Test_Program")),
                            DateStr = ConvertToDateString(reader.GetValue(reader.GetOrdinal("OpenDate"))),
                            Wafer_ID = ConvertToString(reader.GetValue(reader.GetOrdinal("Wafer_ID"))),
                            Wafer_Total_Dies = ConvertToInt32(reader.GetValue(reader.GetOrdinal("Wafer_Total_Dies"))),
                            SBIN = ConvertToString(reader.GetValue(reader.GetOrdinal("SBIN"))),
                            Card_ID = ConvertToString(reader.GetValue(reader.GetOrdinal("CARD_ID")))
                        };

                        data.Add(record);
                    }
                }
            }
        }

        return data;
    }

    /// <summary>
    /// 第一步：查询所有设备（参考Delphi代码第11-39行）
    /// </summary>
    public async Task<List<string>> QueryDevicesAsync(DateTime startDate, DateTime endDate)
    {
        var devices = new List<string>();

        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = @"
                SELECT device
                FROM rtm_admin.rtm_p_data_head
                WHERE lot_id LIKE 'GMT%'
                    AND ht_code = 'BB30'
                    AND status = 'FINISH'
                    AND close_date >= :startDate
                    AND close_date <= :endDate
                GROUP BY device
                ORDER BY device";

            using (var command = new OracleCommand(sql, connection))
            {
                command.Parameters.Add(":startDate", OracleDbType.Varchar2).Value = startDate.ToString("yyyy-MM-dd 00:00:00");
                command.Parameters.Add(":endDate", OracleDbType.Varchar2).Value = endDate.ToString("yyyy-MM-dd 23:59:59");

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        devices.Add(reader.GetString(0));
                    }
                }
            }
        }

        return devices;
    }

    /// <summary>
    /// 第二步：查询设备和CP的组合（参考Delphi代码第46-63行）
    /// </summary>
    public async Task<List<DeviceCPCombination>> QueryDeviceCPCombinationsAsync(DateTime startDate, DateTime endDate)
    {
        var combinations = new List<DeviceCPCombination>();

        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = @"
                SELECT SUBSTR(device, 1, 6) as device, cp_no
                FROM rtm_admin.rtm_p_data_head
                WHERE lot_id LIKE 'GMT%'
                    AND ht_code = 'BB30'
                    AND status = 'FINISH'
                    AND close_date >= :startDate
                    AND close_date <= :endDate
                GROUP BY SUBSTR(device, 1, 6), cp_no
                ORDER BY SUBSTR(device, 1, 6), cp_no";

            using (var command = new OracleCommand(sql, connection))
            {
                command.Parameters.Add(":startDate", OracleDbType.Varchar2).Value = startDate.ToString("yyyy-MM-dd 00:00:00");
                command.Parameters.Add(":endDate", OracleDbType.Varchar2).Value = endDate.ToString("yyyy-MM-dd 23:59:59");

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        combinations.Add(new DeviceCPCombination
                        {
                            Device = reader.GetString(0),
                            CP_No = reader.GetString(1)
                        });
                    }
                }
            }
        }

        return combinations;
    }

    /// <summary>
    /// 第三步：查询详细数据（参考Delphi代码第79-97行）
    /// </summary>
    public async Task<List<GMTDetailRecord>> QueryDetailDataAsync(DateTime startDate, DateTime endDate, string device, string cpNo)
    {
        var details = new List<GMTDetailRecord>();

        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = @"
                SELECT device, wf_lot, serial_no, cp_no, rp_no, tested_wf_qty
                FROM rtm_admin.rtm_p_data_head
                WHERE lot_id LIKE 'GMT%'
                    AND ht_code = 'BB30'
                    AND status = 'FINISH'
                    AND close_date >= :startDate
                    AND close_date <= :endDate
                    AND cp_no = :cpNo
                    AND device LIKE :devicePattern
                GROUP BY serial_no, device, wf_lot, cp_no, rp_no, tested_wf_qty
                ORDER BY serial_no, wf_lot, cp_no, rp_no";

            using (var command = new OracleCommand(sql, connection))
            {
                command.Parameters.Add(":startDate", OracleDbType.Varchar2).Value = startDate.ToString("yyyy-MM-dd 00:00:00");
                command.Parameters.Add(":endDate", OracleDbType.Varchar2).Value = endDate.ToString("yyyy-MM-dd 23:59:59");
                command.Parameters.Add(":cpNo", OracleDbType.Varchar2).Value = cpNo;
                command.Parameters.Add(":devicePattern", OracleDbType.Varchar2).Value = device + "%";

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        details.Add(new GMTDetailRecord
                        {
                            Device = reader.GetString(reader.GetOrdinal("device")),
                            WF_Lot = reader.GetString(reader.GetOrdinal("wf_lot")),
                            Serial_No = reader.GetString(reader.GetOrdinal("serial_no")),
                            CP_No = reader.GetString(reader.GetOrdinal("cp_no")),
                            RP_No = reader.GetString(reader.GetOrdinal("rp_no")),
                            Tested_WF_Qty = ConvertToInt32(reader.GetValue(reader.GetOrdinal("tested_wf_qty")))
                        });
                    }
                }
            }
        }

        return details;
    }

    /// <summary>
    /// 查询P_DATA_SPLIT获取TEST_P（参考Delphi代码第126-133行）
    /// </summary>
    public async Task<string> QueryTestProgramAsync(string serialNo)
    {
        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = "SELECT test_p FROM P_data_split WHERE serial_no = :serialNo";

            using (var command = new OracleCommand(sql, connection))
            {
                command.Parameters.Add(":serialNo", OracleDbType.Varchar2).Value = serialNo;

                var result = await command.ExecuteScalarAsync();
                return result?.ToString() ?? string.Empty;
            }
        }
    }

    /// <summary>
    /// 查询P_DATA_DETAIL获取测试详情（参考Delphi代码第135-140行）
    /// </summary>
    public async Task<List<GMTTestDetail>> QueryTestDetailsAsync(string serialNo)
    {
        var details = new List<GMTTestDetail>();

        using (var connection = new OracleConnection(_connectionString))
        {
            await connection.OpenAsync();

            var sql = @"
                SELECT wf_no, tester, start_date, card_id, sbin, hbin
                FROM P_data_detail
                WHERE serial_no = :serialNo
                ORDER BY TO_NUMBER(wf_no)";

            using (var command = new OracleCommand(sql, connection))
            {
                command.Parameters.Add(":serialNo", OracleDbType.Varchar2).Value = serialNo;

                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        var ordStart = reader.GetOrdinal("start_date");
                        var ordHbin = reader.GetOrdinal("hbin");
                        var ordCard = reader.GetOrdinal("card_id");
                        var startDate = reader.IsDBNull(ordStart)
                            ? DateTime.MinValue
                            : ConvertOracleValueToDateTime(reader.GetValue(ordStart));

                        details.Add(new GMTTestDetail
                        {
                            WF_No = reader.GetString(reader.GetOrdinal("wf_no")),
                            Tester = reader.GetString(reader.GetOrdinal("tester")),
                            Start_Date = startDate,
                            Card_ID = reader.IsDBNull(ordCard) ? string.Empty : reader.GetString(ordCard),
                            SBIN = reader.IsDBNull(reader.GetOrdinal("sbin")) ? string.Empty : reader.GetString(reader.GetOrdinal("sbin")),
                            HBIN = reader.IsDBNull(ordHbin) ? string.Empty : reader.GetString(ordHbin)
                        });
                    }
                }
            }
        }

        return details;
    }
}
