using System.Diagnostics;

namespace IpTesterEnhanced;

class Program
{
    static async Task<int> Main(string[] args)
    {
        Console.Title = "IP连通性测试 - 增强版";
        Console.WriteLine("=== IP连通性测试 - 基于C# ===");
        Console.WriteLine();

        try
        {
            // 初始化配置
            var config = new ConfigManager();
            config.DisplayConfig();
            Console.WriteLine();

            // 查找Excel文件
            var excelFiles = FindExcelFiles();
            if (excelFiles.Length == 0)
            {
                Console.WriteLine("错误：在当前目录下没有找到任何 .xlsx 文件。");
                Console.WriteLine("请将要测试的Excel文件放在程序目录下。");
                WaitForExit();
                return 1;
            }

            Console.WriteLine($"找到 {excelFiles.Length} 个Excel文件:");
            foreach (var file in excelFiles)
            {
                Console.WriteLine($"  - {Path.GetFileName(file)}");
            }
            Console.WriteLine();

            // 创建取消令牌
            using var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (_, e) =>
            {
                e.Cancel = true;
                cts.Cancel();
                Console.WriteLine("\n正在取消操作...");
            };

            // 执行测试
            var success = await RunTestAsync(config, excelFiles, cts.Token);
            
            Console.WriteLine("\n所有操作完成。");
            WaitForExit();
            return success ? 0 : 1;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("\n操作已被用户取消。");
            WaitForExit();
            return 2;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n发生未知错误: {ex.Message}");
            Console.WriteLine($"详细信息: {ex}");
            WaitForExit();
            return 3;
        }
    }

    /// <summary>
    /// 执行IP测试
    /// </summary>
    private static async Task<bool> RunTestAsync(
        ConfigManager config, 
        string[] excelFiles, 
        CancellationToken cancellationToken)
    {
        var totalStopwatch = Stopwatch.StartNew();

        try
        {
            // 1. 读取IP地址
            Console.WriteLine("步骤 1/3: 从Excel文件中读取IP地址...");
            using var excelProcessor = new ExcelProcessor(config);
            var allIps = await excelProcessor.ReadIpsFromExcelFilesAsync(excelFiles);

            if (allIps.Count == 0)
            {
                Console.WriteLine("在指定的Excel文件中未找到有效的IP地址。");
                return false;
            }

            Console.WriteLine($"共找到 {allIps.Count} 个唯一IP地址。");
            Console.WriteLine();

            // 2. 执行Ping测试
            Console.WriteLine("步骤 2/3: 执行IP连通性测试...");
            var progressReporter = new ProgressReporter(config.EnableProgress);
            using var pingTester = new IpPingTester(config, progressReporter);
            
            var results = await pingTester.TestAllIpsAsync(allIps, cancellationToken);
            progressReporter.Complete();

            if (cancellationToken.IsCancellationRequested)
            {
                Console.WriteLine("测试被取消。");
                return false;
            }

            // 3. 写回结果
            Console.WriteLine("步骤 3/3: 将结果写回Excel文件...");
            await excelProcessor.WriteResultsToExcelFilesAsync(excelFiles, results);

            totalStopwatch.Stop();

            // 显示统计信息
            var statistics = TestStatistics.Calculate(
                results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value), 
                totalStopwatch.Elapsed);
            statistics.Display();

            // 保存统计报告
            var reportPath = $"test_report_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            await statistics.SaveToFileAsync(reportPath);

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"测试过程中发生错误: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 查找Excel文件
    /// </summary>
    private static string[] FindExcelFiles()
    {
        var currentDir = Directory.GetCurrentDirectory();
        return Directory.GetFiles(currentDir, "*.xlsx", SearchOption.TopDirectoryOnly)
                       .Where(file => !Path.GetFileName(file).StartsWith("~")) // 排除临时文件
                       .ToArray();
    }

    /// <summary>
    /// 等待用户按键退出
    /// </summary>
    private static void WaitForExit()
    {
        Console.WriteLine();
        Console.WriteLine("按任意键退出...");
        Console.ReadKey(true);
    }
}
