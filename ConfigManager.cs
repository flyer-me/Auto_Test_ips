using System.Configuration;
using System.Text;

namespace IpTesterEnhanced;

/// <summary>
/// 配置管理类，负责读取和管理应用程序配置
/// </summary>
public class ConfigManager
{
    private readonly Dictionary<string, string> _config = new();
    
    public ConfigManager(string configPath = "config.ini")
    {
        LoadConfig(configPath);
    }

    /// <summary>
    /// 从配置文件加载配置
    /// </summary>
    private void LoadConfig(string configPath)
    {
        try
        {
            if (!File.Exists(configPath))
            {
                Console.WriteLine($"警告: 配置文件 {configPath} 不存在，将使用默认配置。");
                SetDefaultConfig();
                return;
            }

            var lines = File.ReadAllLines(configPath, Encoding.UTF8);
            string? currentSection = null;

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                
                // 跳过空行和注释
                if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith('#'))
                    continue;

                // 处理节
                if (trimmedLine.StartsWith('[') && trimmedLine.EndsWith(']'))
                {
                    currentSection = trimmedLine[1..^1];
                    continue;
                }

                // 处理键值对
                var equalIndex = trimmedLine.IndexOf('=');
                if (equalIndex > 0 && currentSection == "cfg")
                {
                    var key = trimmedLine[..equalIndex].Trim();
                    var value = trimmedLine[(equalIndex + 1)..].Trim();
                    _config[key] = value;
                }
            }

            SetDefaultConfig(); // 确保所有必需的配置都有默认值
        }
        catch (Exception ex)
        {
            Console.WriteLine($"读取配置文件时出错: {ex.Message}，将使用默认配置。");
            SetDefaultConfig();
        }
    }

    /// <summary>
    /// 设置默认配置值
    /// </summary>
    private void SetDefaultConfig()
    {
        _config.TryAdd("ping_timeout", "2000");
        _config.TryAdd("ping_times", "3");
        _config.TryAdd("max_concurrent", "100");
        _config.TryAdd("batch_size", "500");
        _config.TryAdd("status_online", "在线");
        _config.TryAdd("status_offline", "离线");
        _config.TryAdd("head_name", "IP");
        _config.TryAdd("enable_progress", "true");
        _config.TryAdd("retry_delay", "100");
    }

    /// <summary>
    /// 获取整数配置值
    /// </summary>
    public int GetInt(string key, int defaultValue = 0)
    {
        if (_config.TryGetValue(key, out var value) && int.TryParse(value, out var result))
            return result;
        return defaultValue;
    }

    /// <summary>
    /// 获取字符串配置值
    /// </summary>
    public string GetString(string key, string defaultValue = "")
    {
        return _config.TryGetValue(key, out var value) ? value : defaultValue;
    }

    /// <summary>
    /// 获取布尔配置值
    /// </summary>
    public bool GetBool(string key, bool defaultValue = false)
    {
        if (_config.TryGetValue(key, out var value))
        {
            return value.ToLowerInvariant() switch
            {
                "true" or "1" or "yes" or "on" => true,
                "false" or "0" or "no" or "off" => false,
                _ => defaultValue
            };
        }
        return defaultValue;
    }

    // 便捷属性
    public int PingTimeout => GetInt("ping_timeout", 2000);
    public int PingTimes => GetInt("ping_times", 3);
    public int MaxConcurrent => GetInt("max_concurrent", 100);
    public int BatchSize => GetInt("batch_size", 500);
    public string StatusOnline => GetString("status_online", "在线");
    public string StatusOffline => GetString("status_offline", "离线");
    public string HeadName => GetString("head_name", "IP");
    public bool EnableProgress => GetBool("enable_progress", true);
    public int RetryDelay => GetInt("retry_delay", 100);

    /// <summary>
    /// 获取IP列名称列表
    /// </summary>
    public string[] IpColumnNames => new[] { HeadName, "专网IP", "IP地址", "IP" };

    /// <summary>
    /// 显示当前配置
    /// </summary>
    public void DisplayConfig()
    {
        Console.WriteLine("当前配置:");
        Console.WriteLine($"  Ping超时: {PingTimeout}ms");
        Console.WriteLine($"  重试次数: {PingTimes}");
        Console.WriteLine($"  最大并发: {MaxConcurrent}");
        Console.WriteLine($"  批处理大小: {BatchSize}");
        Console.WriteLine($"  在线状态: {StatusOnline}");
        Console.WriteLine($"  离线状态: {StatusOffline}");
        Console.WriteLine($"  显示进度: {EnableProgress}");
        Console.WriteLine($"  重试延迟: {RetryDelay}ms");
    }
}
