# IP连通性测试工具

基于C#开发的IP连通性测试工具，支持从Excel文件读取IP地址，执行并发ping测试，并将结果写回Excel文件。

## 功能特性

- 异步并发ping测试，支持大批量IP处理
- 从Excel文件读取IP地址，自动识别IP列
- 测试结果写回Excel，用颜色标识在线/离线状态
- 实时进度显示和详细统计报告
- 可配置的超时时间、重试次数、并发数等参数

## 快速使用

### 1. 编译运行
```bash
dotnet build -c Release
dotnet run
```

### 2. 准备Excel文件
将包含IP地址的Excel文件(.xlsx)放在程序目录下，程序会自动识别IP列。

### 3. 配置参数（可选）
编辑`config.ini`调整测试参数：
```ini
[cfg]
ping_timeout=2000      # ping超时时间(ms)
ping_times=3           # 重试次数
max_concurrent=100     # 最大并发数
status_online=在线     # 在线状态文本
status_offline=离线    # 离线状态文本
```

## 系统要求

- .NET 8.0 或更高版本
- Windows 10/11
- 支持ICMP协议的网络环境
