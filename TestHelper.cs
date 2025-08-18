using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace IpTesterEnhanced;

/// <summary>
/// 测试辅助工具类
/// </summary>
public static class TestHelper
{
    /// <summary>
    /// 创建测试用的Excel文件
    /// </summary>
    public static async Task CreateTestExcelFileAsync(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("IP测试");
        
        // 添加标题行
        worksheet.Cells[1, 1].Value = "序号";
        worksheet.Cells[1, 2].Value = "IP地址";
        worksheet.Cells[1, 3].Value = "描述";
        
        // 添加测试数据
        var testIps = new[]
        {
            ("8.8.8.8", "Google DNS"),
            ("8.8.4.4", "Google DNS 备用"),
            ("114.114.114.114", "114 DNS"),
            ("223.5.5.5", "阿里 DNS"),
            ("119.29.29.29", "腾讯 DNS"),
            ("1.1.1.1", "Cloudflare DNS"),
            ("192.168.1.1", "常见路由器IP"),
            ("10.0.0.1", "内网IP"),
            ("172.16.0.1", "内网IP"),
            ("192.168.0.1", "常见路由器IP")
        };
        
        for (int i = 0; i < testIps.Length; i++)
        {
            worksheet.Cells[i + 2, 1].Value = i + 1;
            worksheet.Cells[i + 2, 2].Value = testIps[i].Item1;
            worksheet.Cells[i + 2, 3].Value = testIps[i].Item2;
        }
        
        // 设置列宽
        worksheet.Column(1).Width = 8;
        worksheet.Column(2).Width = 15;
        worksheet.Column(3).Width = 20;
        
        // 设置标题行样式
        using var headerRange = worksheet.Cells[1, 1, 1, 3];
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
        
        await package.SaveAsAsync(new FileInfo(filePath));
        Console.WriteLine($"测试Excel文件已创建: {filePath}");
    }
    
    /// <summary>
    /// 验证程序输出
    /// </summary>
    public static void ValidateOutput(string excelPath)
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var worksheet = package.Workbook.Worksheets.First();
            
            // 查找状态列
            int statusColumn = -1;
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var header = worksheet.Cells[1, col].Value?.ToString();
                if (header != null && (header.Contains("状态") || header.Contains("结果")))
                {
                    statusColumn = col;
                    break;
                }
            }
            
            if (statusColumn == -1)
            {
                Console.WriteLine("❌ 未找到状态列");
                return;
            }
            
            Console.WriteLine("✅ 找到状态列");
            
            // 检查结果
            int onlineCount = 0;
            int offlineCount = 0;
            
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var status = worksheet.Cells[row, statusColumn].Value?.ToString();
                if (status == "在线") onlineCount++;
                else if (status == "离线") offlineCount++;
            }
            
            Console.WriteLine($"✅ 测试结果统计: 在线 {onlineCount} 个, 离线 {offlineCount} 个");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 验证输出时出错: {ex.Message}");
        }
    }
}
