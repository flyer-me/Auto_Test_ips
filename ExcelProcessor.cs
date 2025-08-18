using System.Collections.Concurrent;
using System.Net;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace IpTesterEnhanced;

/// <summary>
/// 异步Excel文件处理器
/// </summary>
public class ExcelProcessor : IDisposable
{
    private readonly ConfigManager _config;
    private readonly SemaphoreSlim _fileSemaphore;

    public ExcelProcessor(ConfigManager config)
    {
        _config = config;
        // 限制同时处理的Excel文件数量，避免内存占用过高
        _fileSemaphore = new SemaphoreSlim(Environment.ProcessorCount, Environment.ProcessorCount);
        
        // 设置EPPlus许可证
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// 异步从多个Excel文件中读取所有IP地址
    /// </summary>
    public async Task<HashSet<string>> ReadIpsFromExcelFilesAsync(string[] filePaths)
    {
        var allIps = new ConcurrentBag<string>();
        
        Console.WriteLine($"开始从 {filePaths.Length} 个Excel文件中读取IP地址...");

        var tasks = filePaths.Select(async filePath =>
        {
            await _fileSemaphore.WaitAsync();
            try
            {
                var ips = await ReadIpsFromSingleFileAsync(filePath);
                foreach (var ip in ips)
                {
                    allIps.Add(ip);
                }
            }
            finally
            {
                _fileSemaphore.Release();
            }
        });

        await Task.WhenAll(tasks);

        var uniqueIps = new HashSet<string>(allIps);
        Console.WriteLine($"共找到 {uniqueIps.Count} 个唯一IP地址");
        
        return uniqueIps;
    }

    /// <summary>
    /// 从单个Excel文件中读取IP地址
    /// </summary>
    private async Task<List<string>> ReadIpsFromSingleFileAsync(string filePath)
    {
        var ips = new List<string>();

        try
        {
            await using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var package = new ExcelPackage(stream);

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var ipColumnIndex = FindIpColumn(worksheet);
                if (ipColumnIndex == -1) continue;

                var dimension = worksheet.Dimension;
                if (dimension == null) continue;

                // 从第二行开始读取（跳过标题行）
                for (int row = dimension.Start.Row + 1; row <= dimension.End.Row; row++)
                {
                    var cellValue = worksheet.Cells[row, ipColumnIndex].Value?.ToString()?.Trim();
                    if (IsValidIpAddress(cellValue))
                    {
                        ips.Add(cellValue!);
                    }
                }
            }

            Console.WriteLine($"从 {Path.GetFileName(filePath)} 读取到 {ips.Count} 个IP地址");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"读取文件 {Path.GetFileName(filePath)} 时出错: {ex.Message}");
        }

        return ips;
    }

    /// <summary>
    /// 异步将测试结果写回Excel文件
    /// </summary>
    public async Task WriteResultsToExcelFilesAsync(
        string[] filePaths, 
        ConcurrentDictionary<string, PingResult> results)
    {
        Console.WriteLine($"开始将结果写回 {filePaths.Length} 个Excel文件...");

        var tasks = filePaths.Select(async filePath =>
        {
            await _fileSemaphore.WaitAsync();
            try
            {
                await WriteResultsToSingleFileAsync(filePath, results);
            }
            finally
            {
                _fileSemaphore.Release();
            }
        });

        await Task.WhenAll(tasks);
        Console.WriteLine("所有结果已成功写回Excel文件");
    }

    /// <summary>
    /// 将结果写回单个Excel文件
    /// </summary>
    private async Task WriteResultsToSingleFileAsync(
        string filePath, 
        ConcurrentDictionary<string, PingResult> results)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var ipColumnIndex = FindIpColumn(worksheet);
                if (ipColumnIndex == -1) continue;

                var statusColumnIndex = GetOrCreateStatusColumn(worksheet);
                var dimension = worksheet.Dimension;
                if (dimension == null) continue;

                // 处理数据行
                for (int row = dimension.Start.Row + 1; row <= dimension.End.Row; row++)
                {
                    var ip = worksheet.Cells[row, ipColumnIndex].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(ip) || !results.TryGetValue(ip, out var result)) 
                        continue;

                    var statusCell = worksheet.Cells[row, statusColumnIndex];
                    
                    // 设置状态文本
                    statusCell.Value = result.IsSuccess ? _config.StatusOnline : _config.StatusOffline;
                    
                    // 设置背景颜色
                    statusCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    statusCell.Style.Fill.BackgroundColor.SetColor(
                        result.IsSuccess ? System.Drawing.Color.LightGreen : System.Drawing.Color.Salmon);
                    
                    // 添加详细信息到注释（可选）
                    if (result.Attempts.Count > 1)
                    {
                        var comment = $"成功率: {result.SuccessRate:P1}\n" +
                                    $"平均延迟: {result.AverageRoundtripTime}ms\n" +
                                    $"测试次数: {result.Attempts.Count}";
                        statusCell.AddComment(comment, "IpTester");
                    }
                }
            }

            await package.SaveAsync();
            Console.WriteLine($"结果已写入 {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"写入文件 {Path.GetFileName(filePath)} 时出错: {ex.Message}");
        }
    }

    /// <summary>
    /// 查找IP列的索引
    /// </summary>
    private int FindIpColumn(ExcelWorksheet worksheet)
    {
        var dimension = worksheet.Dimension;
        if (dimension == null) return -1;

        // 检查前5行以查找IP列标题
        for (int row = dimension.Start.Row; row <= Math.Min(dimension.Start.Row + 4, dimension.End.Row); row++)
        {
            for (int col = dimension.Start.Column; col <= dimension.End.Column; col++)
            {
                var header = worksheet.Cells[row, col].Value?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(header) && 
                    _config.IpColumnNames.Any(name => 
                        string.Equals(header, name, StringComparison.OrdinalIgnoreCase)))
                {
                    return col;
                }
            }
        }

        return -1;
    }

    /// <summary>
    /// 获取或创建状态列
    /// </summary>
    private int GetOrCreateStatusColumn(ExcelWorksheet worksheet)
    {
        var dimension = worksheet.Dimension;
        if (dimension == null)
        {
            worksheet.Cells[1, 1].Value = "状态";
            return 1;
        }

        // 查找现有的状态列
        for (int col = dimension.Start.Column; col <= dimension.End.Column; col++)
        {
            var header = worksheet.Cells[1, col].Value?.ToString()?.Trim();
            if (!string.IsNullOrEmpty(header) && 
                (header.Contains("状态") || header.Contains("结果") || header.Contains("Status")))
            {
                return col;
            }
        }

        // 创建新的状态列
        var newColIndex = dimension.End.Column + 1;
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        worksheet.Cells[1, newColIndex].Value = $"测试结果_{timestamp}";
        
        // 设置标题样式
        var headerCell = worksheet.Cells[1, newColIndex];
        headerCell.Style.Font.Bold = true;
        headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        headerCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
        
        return newColIndex;
    }

    /// <summary>
    /// 验证IP地址格式
    /// </summary>
    private static bool IsValidIpAddress(string? ipString)
    {
        return !string.IsNullOrWhiteSpace(ipString) && IPAddress.TryParse(ipString, out _);
    }

    public void Dispose()
    {
        _fileSemaphore?.Dispose();
    }
}
