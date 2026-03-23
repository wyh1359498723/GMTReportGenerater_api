var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// 注册GMT服务
builder.Services.AddSingleton<GMTReportGenerater_api.Services.GMTConfigService>();
builder.Services.AddScoped<GMTReportGenerater_api.Services.GMTDataProcessingService>();
builder.Services.AddScoped<GMTReportGenerater_api.Services.GMTDatabaseService>();
builder.Services.AddScoped<GMTReportGenerater_api.Services.GMTExcelGenerationService>();
builder.Services.AddScoped<GMTReportGenerater_api.Services.GMTExcelGenerationSpecialService>();
builder.Services.AddScoped<GMTReportGenerater_api.Services.GMTExcelGenerationQcService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
// 显示详细错误信息（便于排查问题）
app.UseDeveloperExceptionPage();

// Swagger在所有环境中都可用（便于测试）
app.UseSwagger();
app.UseSwaggerUI();

// HTTPS重定向仅在开发环境启用（IIS会处理生产环境的HTTPS）
if (app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}

app.UseAuthorization();

app.MapControllers();

app.Run();
