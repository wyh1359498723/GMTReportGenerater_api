# GMT月报生成API

这是一个基于ASP.NET Core的Web API，用于生成GMT客户的月报数据。该API参考了原有的Delphi代码逻辑，提供了RESTful接口来获取结构化的月报数据。

## 功能特性

- 根据时间范围查询GMT测试数据
- 自动解析setup.ini配置文件
- 按设备和CP分组处理数据
- 计算测试良率（Yield）
- 提供结构化的JSON响应
- **生成Excel月报文件**

## 技术栈

- ASP.NET Core 8.0
- Oracle数据库
- EPPlus 7.0 (Excel文件生成)
- Swagger/OpenAPI文档

## 配置

### 数据库连接

在 `appsettings.json` 中配置Oracle数据库连接字符串：

```json
{
  "ConnectionStrings": {
    "OracleDb": "User Id=your_username;Password=your_password;Data Source=your_oracle_connection_string"
  }
}
```

### 配置文件

- `Template/setup.ini`: 包含设备配置和CP参数
- `Template/GMT-BB30_template.xlsx`: Excel模板文件（用于生成月报）
- `Output/`: Excel文件输出目录（自动创建）

## API接口

### 1. 生成GMT月报

**POST** `/api/GMTReport/generate`

时间参数格式：`yyyy-MM-dd HH:mm:ss`（字符串格式）

请求体：
```json
{
  "startDate": "2025-12-31 00:00:00",
  "endDate": "2026-01-21 23:59:59"
}
```

响应示例：
```json
{
  "deviceSummaries": [
    {
      "deviceName": "NS3602",
      "cpData": [
        {
          "cpNumber": "CP1",
          "dataRows": [
            {
              "date": "2025-12-31",
              "tester": "TESTER_1",
              "wfLot": "GMT_TEST_LOT_001",
              "temp": "25",
              "prgName": "PROGRAM_1",
              "wfNo": "1",
              "grossQty": 19873,
              "binCounts": {
                "2": 150,
                "4": 200
              },
              "yield": "85.5%"
            }
          ],
          "totalRow": {
            "yield": "85.5%"
          }
        }
      ]
    }
  ],
  "reportDate": "2025-02-03",
  "dateRange": "2025-12-31 至 2026-01-21"
}
```

### 2. 生成并下载Excel月报文件

**POST** `/api/GMTReport/generate-excel`

时间参数格式：`yyyy-MM-dd HH:mm:ss`（字符串格式）

请求体：
```json
{
  "startDate": "2025-12-31 00:00:00",
  "endDate": "2026-01-21 23:59:59"
}
```

响应：直接下载Excel文件（.xlsx格式）

生成的文件位置：`Output/GMT_Report_YYYYMMDD_YYYYMMDD_YYYYMMDDHHmmss.xlsx`

### 3. 生成并下载Excel月报文件 - Special格式

**POST** `/api/GMTReport/generate-excel-special`

时间参数格式：`yyyy-MM-dd HH:mm:ss`（字符串格式）

请求体：
```json
{
  "startDate": "2025-12-31 00:00:00",
  "endDate": "2026-01-21 23:59:59"
}
```

响应：直接下载Excel文件（.xlsx格式，Special格式）

生成的文件位置：`Output/GMT_Report_Special_YYYYMMDD_YYYYMMDD_YYYYMMDDHHmmss.xlsx`

**Special格式特点**：
- 列顺序：Date, HTSH, Device, CP_NO, WF_LOT, WF_NO, GROSS_QTY, PASS_QTY, FAIL_QTY, YIELD, BIN数据, TESTER, CARD_ID
- 包含完整Device名称（不截取前6位）
- 包含Pass/Fail数量统计
- 模板文件：`GMT-BB30-SPECIAL_template.xlsx`

### 4. 获取设备配置

**GET** `/api/GMTReport/devices`

返回GMT设备列表和配置信息。

## 数据处理逻辑

1. **配置解析**: 从 `setup.ini` 读取设备和CP配置
2. **数据查询**: 根据时间范围查询数据库
3. **数据分组**: 按设备和CP分组处理
4. **Yield计算**: 根据合格bin计算良率
5. **汇总计算**: 生成总计行数据

## 运行项目

```bash
# 还原依赖
dotnet restore

# 运行项目
dotnet run
```

访问 Swagger UI: `https://localhost:5001/swagger`

## 配置说明

### setup.ini 格式

```ini
[GMT]
DEVICE=NS3602,NS3605,NS3606,NS3607,NS3613,NS2001,NS3608,NS3611
DEVICE_ALL=NS3607_2A

[NS3602]
CP1=19873*$2$&AX|2,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,
```

其中CP配置格式为：`总数*$合格bin$&单元格边界|显示bin列表`

## 注意事项

- 确保数据库连接字符串正确配置
- 时间范围不能超过一年
- 需要确保Oracle客户端库可用
- 配置文件路径为 `Template/setup.ini`
