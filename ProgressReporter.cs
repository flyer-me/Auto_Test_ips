using System.Diagnostics;
using System.Text;

namespace IpTesterEnhanced;

/// <summary>
/// 进度报告器，提供实时进度显示和统计功能
/// </summary>
public class ProgressReporter : IProgress<PingProgress>
{
    private readonly bool _enableProgress;
    private readonly object _lockObject = new();
    private int _lastProgressLength = 0;
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();

    public ProgressReporter(bool enableProgress = true)
    {
        _enableProgress = enableProgress;
    }

    /// <summary>
    /// 报告进度更新
    /// </summary>
    public void Report(PingProgress progress)
    {
        if (!_enableProgress) return;

        lock (_lockObject)
        {
            // 清除之前的进度行
            if (_lastProgressLength > 0)
            {
                Console.SetCursorPosition(0, Console.CursorTop);
                Console.Write(new string(' ', _lastProgressLength));
                Console.SetCursorPosition(0, Console.CursorTop);
            }

            // 构建进度信息
            var progressText = BuildProgressText(progress);
            Console.Write(progressText);
            _lastProgressLength = progressText.Length;
        }
    }

    /// <summary>
    /// 构建进度文本
    /// </summary>
    private string BuildProgressText(PingProgress progress)
    {
        var sb = new StringBuilder();
        
        // 基本进度信息
        sb.Append($"进度: {progress.CompletedCount}/{progress.TotalCount} ");
        sb.Append($"({progress.ProgressPercentage:F1}%) ");
        
        // 进度条
        var progressBarWidth = 20;
        var filledWidth = (int)(progress.ProgressPercentage / 100 * progressBarWidth);
        sb.Append('[');
        sb.Append(new string('█', filledWidth));
        sb.Append(new string('░', progressBarWidth - filledWidth));
        sb.Append("] ");
        
        // 时间信息
        sb.Append($"已用时: {FormatTimeSpan(progress.ElapsedTime)} ");
        
        if (progress.CompletedCount > 0)
        {
            var estimatedRemaining = TimeSpan.FromSeconds(progress.EstimatedTimeRemaining);
            sb.Append($"预计剩余: {FormatTimeSpan(estimatedRemaining)} ");
            
            // 速度信息
            var speed = progress.CompletedCount / progress.ElapsedTime.TotalSeconds;
            sb.Append($"速度: {speed:F1} IP/秒");
        }

        return sb.ToString();
    }

    /// <summary>
    /// 格式化时间跨度
    /// </summary>
    private static string FormatTimeSpan(TimeSpan timeSpan)
    {
        if (timeSpan.TotalHours >= 1)
            return $"{timeSpan.Hours:D2}:{timeSpan.Minutes:D2}:{timeSpan.Seconds:D2}";
        else
            return $"{timeSpan.Minutes:D2}:{timeSpan.Seconds:D2}";
    }

    /// <summary>
    /// 完成进度报告
    /// </summary>
    public void Complete()
    {
        if (!_enableProgress) return;

        lock (_lockObject)
        {
            Console.WriteLine(); // 换行
        }
    }
}

/// <summary>
/// 测试统计信息
/// </summary>
public class TestStatistics
{
    public int TotalIpCount { get; set; }
    public int OnlineCount { get; set; }
    public int OfflineCount { get; set; }
    public TimeSpan TotalTestTime { get; set; }
    public double AverageResponseTime { get; set; }
    public double TestSpeed { get; set; }
    public Dictionary<string, int> StatusDistribution { get; set; } = new();

    /// <summary>
    /// 计算统计信息
    /// </summary>
    public static TestStatistics Calculate(
        Dictionary<string, PingResult> results, 
        TimeSpan totalTime)
    {
        var stats = new TestStatistics
        {
            TotalIpCount = results.Count,
            TotalTestTime = totalTime
        };

        var totalResponseTime = 0L;
        var responseTimeCount = 0;

        foreach (var result in results.Values)
        {
            if (result.IsSuccess)
            {
                stats.OnlineCount++;
                totalResponseTime += result.AverageRoundtripTime;
                responseTimeCount++;
            }
            else
            {
                stats.OfflineCount++;
            }

            // 统计状态分布
            var status = result.IsSuccess ? "在线" : "离线";
            stats.StatusDistribution[status] = stats.StatusDistribution.GetValueOrDefault(status, 0) + 1;
        }

        stats.AverageResponseTime = responseTimeCount > 0 ? (double)totalResponseTime / responseTimeCount : 0;
        stats.TestSpeed = totalTime.TotalSeconds > 0 ? stats.TotalIpCount / totalTime.TotalSeconds : 0;

        return stats;
    }

    /// <summary>
    /// 显示统计信息
    /// </summary>
    public void Display()
    {
        Console.WriteLine("\n=== 测试统计 ===");
        Console.WriteLine($"总IP数量: {TotalIpCount}");
        Console.WriteLine($"在线数量: {OnlineCount} ({(double)OnlineCount / TotalIpCount:P1})");
        Console.WriteLine($"离线数量: {OfflineCount} ({(double)OfflineCount / TotalIpCount:P1})");
        Console.WriteLine($"总测试时间: {FormatTimeSpan(TotalTestTime)}");
        Console.WriteLine($"平均响应时间: {AverageResponseTime:F1} ms");
        Console.WriteLine($"测试速度: {TestSpeed:F1} IP/秒");
        
        if (StatusDistribution.Count > 0)
        {
            Console.WriteLine("\n状态分布:");
            foreach (var kvp in StatusDistribution)
            {
                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
            }
        }
        
        Console.WriteLine("================");
    }

    /// <summary>
    /// 格式化时间跨度
    /// </summary>
    private static string FormatTimeSpan(TimeSpan timeSpan)
    {
        if (timeSpan.TotalHours >= 1)
            return $"{timeSpan.Hours:D2}:{timeSpan.Minutes:D2}:{timeSpan.Seconds:D2}";
        else
            return $"{timeSpan.Minutes:D2}:{timeSpan.Seconds:D2}";
    }

    /// <summary>
    /// 保存统计信息到文件
    /// </summary>
    public async Task SaveToFileAsync(string filePath)
    {
        var content = new StringBuilder();
        content.AppendLine("IP连通性测试统计报告");
        content.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        content.AppendLine();
        content.AppendLine($"总IP数量: {TotalIpCount}");
        content.AppendLine($"在线数量: {OnlineCount} ({(double)OnlineCount / TotalIpCount:P1})");
        content.AppendLine($"离线数量: {OfflineCount} ({(double)OfflineCount / TotalIpCount:P1})");
        content.AppendLine($"总测试时间: {FormatTimeSpan(TotalTestTime)}");
        content.AppendLine($"平均响应时间: {AverageResponseTime:F1} ms");
        content.AppendLine($"测试速度: {TestSpeed:F1} IP/秒");
        
        if (StatusDistribution.Count > 0)
        {
            content.AppendLine();
            content.AppendLine("状态分布:");
            foreach (var kvp in StatusDistribution)
            {
                content.AppendLine($"  {kvp.Key}: {kvp.Value}");
            }
        }

        await File.WriteAllTextAsync(filePath, content.ToString(), Encoding.UTF8);
        Console.WriteLine($"统计报告已保存到: {filePath}");
    }
}
