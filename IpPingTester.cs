using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;

namespace IpTesterEnhanced;

/// <summary>
/// 高性能异步IP连通性测试器
/// </summary>
public class IpPingTester : IDisposable
{
    private readonly ConfigManager _config;
    private readonly SemaphoreSlim _semaphore;
    private readonly IProgress<PingProgress>? _progress;

    public IpPingTester(ConfigManager config, IProgress<PingProgress>? progress = null)
    {
        _config = config;
        _semaphore = new SemaphoreSlim(_config.MaxConcurrent, _config.MaxConcurrent);
        _progress = progress;
    }

    /// <summary>
    /// 异步测试所有IP地址的连通性
    /// </summary>
    public async Task<ConcurrentDictionary<string, PingResult>> TestAllIpsAsync(
        IEnumerable<string> ips, 
        CancellationToken cancellationToken = default)
    {
        var ipList = ips.ToList();
        var results = new ConcurrentDictionary<string, PingResult>();
        var totalCount = ipList.Count;
        var completedCount = 0;

        Console.WriteLine($"开始测试 {totalCount} 个IP地址...");
        var stopwatch = Stopwatch.StartNew();

        // 分批处理IP地址以提高性能
        var batches = ipList.Chunk(_config.BatchSize);
        
        foreach (var batch in batches)
        {
            if (cancellationToken.IsCancellationRequested)
                break;

            var batchTasks = batch.Select(ip => TestSingleIpAsync(ip, results, cancellationToken))
                                  .ToArray();

            await Task.WhenAll(batchTasks);

            // 更新进度
            var batchCompleted = Interlocked.Add(ref completedCount, batch.Length);
            _progress?.Report(new PingProgress
            {
                CompletedCount = batchCompleted,
                TotalCount = totalCount,
                ElapsedTime = stopwatch.Elapsed
            });
        }

        stopwatch.Stop();
        Console.WriteLine($"测试完成，总耗时: {stopwatch.Elapsed.TotalSeconds:F2} 秒");
        
        return results;
    }

    /// <summary>
    /// 测试单个IP地址
    /// </summary>
    private async Task TestSingleIpAsync(
        string ip, 
        ConcurrentDictionary<string, PingResult> results,
        CancellationToken cancellationToken)
    {
        await _semaphore.WaitAsync(cancellationToken);
        
        try
        {
            var result = await PingWithRetryAsync(ip, cancellationToken);
            results[ip] = result;
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// 带重试机制的Ping测试
    /// </summary>
    private async Task<PingResult> PingWithRetryAsync(string ip, CancellationToken cancellationToken)
    {
        var attempts = new List<PingAttempt>();
        var isSuccess = false;
        long totalRoundtripTime = 0;
        var successCount = 0;

        for (int attempt = 0; attempt < _config.PingTimes; attempt++)
        {
            if (cancellationToken.IsCancellationRequested)
                break;

            try
            {
                using var ping = new Ping();
                var reply = await ping.SendPingAsync(ip, _config.PingTimeout);
                
                var attemptResult = new PingAttempt
                {
                    AttemptNumber = attempt + 1,
                    Status = reply.Status,
                    RoundtripTime = reply.RoundtripTime,
                    Timestamp = DateTime.Now
                };

                attempts.Add(attemptResult);

                if (reply.Status == IPStatus.Success)
                {
                    isSuccess = true;
                    totalRoundtripTime += reply.RoundtripTime;
                    successCount++;
                    
                    // 如果第一次就成功，可以提前结束（可选优化）
                    if (attempt == 0)
                        break;
                }
                else if (attempt < _config.PingTimes - 1)
                {
                    // 在重试之间添加短暂延迟
                    await Task.Delay(_config.RetryDelay, cancellationToken);
                }
            }
            catch (PingException ex)
            {
                attempts.Add(new PingAttempt
                {
                    AttemptNumber = attempt + 1,
                    Status = IPStatus.Unknown,
                    RoundtripTime = 0,
                    Timestamp = DateTime.Now,
                    ErrorMessage = ex.Message
                });
            }
            catch (Exception ex)
            {
                attempts.Add(new PingAttempt
                {
                    AttemptNumber = attempt + 1,
                    Status = IPStatus.Unknown,
                    RoundtripTime = 0,
                    Timestamp = DateTime.Now,
                    ErrorMessage = ex.Message
                });
            }
        }

        return new PingResult
        {
            IpAddress = ip,
            IsSuccess = isSuccess,
            AverageRoundtripTime = successCount > 0 ? totalRoundtripTime / successCount : 0,
            SuccessRate = (double)successCount / attempts.Count,
            Attempts = attempts,
            TestTimestamp = DateTime.Now
        };
    }

    public void Dispose()
    {
        _semaphore?.Dispose();
    }
}

/// <summary>
/// Ping测试结果
/// </summary>
public class PingResult
{
    public string IpAddress { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public long AverageRoundtripTime { get; set; }
    public double SuccessRate { get; set; }
    public List<PingAttempt> Attempts { get; set; } = new();
    public DateTime TestTimestamp { get; set; }
}

/// <summary>
/// 单次Ping尝试结果
/// </summary>
public class PingAttempt
{
    public int AttemptNumber { get; set; }
    public IPStatus Status { get; set; }
    public long RoundtripTime { get; set; }
    public DateTime Timestamp { get; set; }
    public string? ErrorMessage { get; set; }
}

/// <summary>
/// 进度报告
/// </summary>
public class PingProgress
{
    public int CompletedCount { get; set; }
    public int TotalCount { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public double ProgressPercentage => TotalCount > 0 ? (double)CompletedCount / TotalCount * 100 : 0;
    public double EstimatedTimeRemaining => CompletedCount > 0 ? 
        ElapsedTime.TotalSeconds * (TotalCount - CompletedCount) / CompletedCount : 0;
}
