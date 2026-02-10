using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using Windows.Devices.Scanners;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Streams;

namespace ScannerCLI
{
    class Program
    {
        private static ImageScanner _scanner;
        private static CancellationTokenSource _cancellationTokenSource;
        private static CommandLineOptions _options = new CommandLineOptions();

        static async Task Main(string[] args)
        {
            try
            {
                // 解析命令行参数
                if (!ParseArguments(args))
                {
                    ShowHelp();
                    return;
                }

                // 如果指定列出设备
                if (_options.ListDevices)
                {
                    await ListScannersAsync();
                    return;
                }

                // 查找扫描仪
                ImageScanner scanner;
                if (!string.IsNullOrEmpty(_options.DeviceName))
                {
                    scanner = await FindScannerByNameAsync(_options.DeviceName);
                    if (scanner == null)
                    {
                        Console.WriteLine($"未找到名为 '{_options.DeviceName}' 的扫描仪设备!");
                        return;
                    }
                }
                else
                {
                    scanner = await FindFirstScannerAsync();
                    if (scanner == null)
                    {
                        Console.WriteLine("未找到WIA兼容扫描仪!");
                        Console.WriteLine("请确保扫描仪已连接并开启。");
                        return;
                    }
                }

                _scanner = scanner;

                // 验证扫描源是否支持
                if (!ValidateScanSource())
                {
                    Console.WriteLine($"扫描仪不支持 '{_options.Source}' 扫描源!");
                    await DisplayScannerInfoAsync();
                    return;
                }

                // 验证分辨率
                if (!ValidateResolution())
                {
                    Console.WriteLine("分辨率必须在 50-1200 DPI 之间!");
                    return;
                }

                // 执行扫描
                await PerformScanAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"内部错误: {ex.InnerException.Message}");
                }
            }
        }

        private static bool ParseArguments(string[] args)
        {
            if (args.Length == 0)
            {
                ShowHelp();
                return false;
            }

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-L":
                    case "--list":
                        _options.ListDevices = true;
                        break;

                    case "-d":
                    case "--device":
                        if (i + 1 < args.Length)
                        {
                            _options.DeviceName = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("错误: -d 参数需要设备名称");
                            return false;
                        }
                        break;

                    case "-r":
                    case "--resolution":
                        if (i + 1 < args.Length)
                        {
                            if (int.TryParse(args[++i], out int resolution))
                            {
                                _options.Resolution = resolution;
                            }
                            else
                            {
                                Console.WriteLine("错误: 分辨率必须是数字");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("错误: -r 参数需要分辨率值");
                            return false;
                        }
                        break;

                    case "-s":
                    case "--source":
                        if (i + 1 < args.Length)
                        {
                            _options.Source = args[++i].ToLower();
                            if (_options.Source != "flatbed" && _options.Source != "feeder")
                            {
                                Console.WriteLine("错误: 扫描源必须是 'flatbed' 或 'feeder'");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("错误: -s 参数需要扫描源");
                            return false;
                        }
                        break;

                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            _options.OutputDirectory = args[++i];
                        }
                        else
                        {
                            Console.WriteLine("错误: -o 参数需要输出目录");
                            return false;
                        }
                        break;

                    case "-f":
                    case "--format":
                        if (i + 1 < args.Length)
                        {
                            _options.Format = args[++i].ToLower();
                            if (!new[] { "pdf", "jpeg", "jpg", "png", "tiff", "bmp" }.Contains(_options.Format))
                            {
                                Console.WriteLine("错误: 不支持的文件格式");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("错误: -f 参数需要文件格式");
                            return false;
                        }
                        break;

                    case "-m":
                    case "--mode":
                        if (i + 1 < args.Length)
                        {
                            _options.ColorMode = args[++i].ToLower();
                            if (!new[] { "lineart", "gray", "color" }.Contains(_options.ColorMode))
                            {
                                Console.WriteLine("错误: 不支持的颜色模式");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("错误: -m 参数需要颜色模式");
                            return false;
                        }
                        break;

                    case "-h":
                    case "--help":
                        ShowHelp();
                        return false;

                    default:
                        Console.WriteLine($"未知参数: {args[i]}");
                        ShowHelp();
                        return false;
                }
            }

            // 验证必要参数（如果不是列出设备）
            if (!_options.ListDevices)
            {
                if (string.IsNullOrEmpty(_options.Source))
                {
                    Console.WriteLine("错误: 必须指定扫描源 (-s)");
                    return false;
                }
            }

            return true;
        }

        private static void ShowHelp()
        {
            Console.WriteLine("扫描仪命令行工具");
            Console.WriteLine("用法: ScannerCLI [选项]");
            Console.WriteLine();
            Console.WriteLine("选项:");
            Console.WriteLine("  -L, --list                列出所有可用的扫描仪设备");
            Console.WriteLine("  -d, --device NAME         通过名称指定要使用的扫描仪设备");
            Console.WriteLine("  -r, --resolution DPI      指定扫描分辨率 (默认: 300)");
            Console.WriteLine("  -s, --source SOURCE       选择扫描源: flatbed 或 feeder");
            Console.WriteLine("  -o, --output DIR          指定输出目录 (默认: 当前目录)");
            Console.WriteLine("  -f, --format FORMAT       指定文件格式: pdf, jpeg, png, tiff, bmp");
            Console.WriteLine("  -m, --mode MODE           选择颜色模式: Lineart, Gray, Color");
            Console.WriteLine("  -h, --help                显示此帮助信息");
            Console.WriteLine();
            Console.WriteLine("示例:");
            Console.WriteLine("  ScannerCLI -L");
            Console.WriteLine("  ScannerCLI -d \"HP Scanner\" -s flatbed -r 300 -f pdf -m Color -o C:\\Scans");
            Console.WriteLine("  ScannerCLI -s feeder -f jpeg -m Gray -r 150");
        }

        private static async Task ListScannersAsync()
        {
            Console.WriteLine("正在查找扫描仪设备...\n");

            try
            {
                var deviceCollection = await DeviceInformation.FindAllAsync(ImageScanner.GetDeviceSelector());

                if (deviceCollection.Count == 0)
                {
                    Console.WriteLine("未找到任何扫描仪设备。");
                    return;
                }

                Console.WriteLine($"找到 {deviceCollection.Count} 个设备:\n");

                for (int i = 0; i < deviceCollection.Count; i++)
                {
                    var device = deviceCollection[i];
                    Console.WriteLine($"[{i + 1}] {device.Name}");
                    Console.WriteLine($"  ID: {device.Id}");

                    try
                    {
                        var scanner = await ImageScanner.FromIdAsync(device.Id);
                        Console.WriteLine("  支持的扫描源:");

                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed))
                            Console.WriteLine("    - Flatbed (平板)");
                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder))
                            Console.WriteLine("    - Feeder (进纸器)");
                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
                            Console.WriteLine("    - Auto (自动)");
                    }
                    catch
                    {
                        Console.WriteLine("  无法获取设备详细信息");
                    }

                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找设备时出错: {ex.Message}");
            }
        }

        private static async Task<ImageScanner> FindFirstScannerAsync()
        {
            try
            {
                var deviceCollection = await DeviceInformation.FindAllAsync(ImageScanner.GetDeviceSelector());

                if (deviceCollection.Count == 0)
                {
                    return null;
                }

                var deviceInfo = deviceCollection[0];
                return await ImageScanner.FromIdAsync(deviceInfo.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找扫描仪时出错: {ex.Message}");
                return null;
            }
        }

        private static async Task<ImageScanner> FindScannerByNameAsync(string deviceName)
        {
            try
            {
                var deviceCollection = await DeviceInformation.FindAllAsync(ImageScanner.GetDeviceSelector());

                foreach (var deviceInfo in deviceCollection)
                {
                    if (deviceInfo.Name.Contains(deviceName, StringComparison.OrdinalIgnoreCase))
                    {
                        return await ImageScanner.FromIdAsync(deviceInfo.Id);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找扫描仪时出错: {ex.Message}");
                return null;
            }
        }

        private static async Task DisplayScannerInfoAsync()
        {
            if (_scanner == null) return;

            Console.WriteLine("\n=== 扫描仪信息 ===");

            var deviceInfo = await DeviceInformation.CreateFromIdAsync(_scanner.DeviceId);
            Console.WriteLine($"设备名称: {deviceInfo.Name}");
            Console.WriteLine($"设备ID: {_scanner.DeviceId}");
            Console.WriteLine("支持的扫描源:");

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed))
                Console.WriteLine("  - Flatbed (平板扫描)");

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder))
            {
                Console.WriteLine("  - Feeder (自动进纸器)");
                var feederConfig = _scanner.FeederConfiguration;
                if (feederConfig != null)
                {
                    Console.WriteLine($"    - 最大页数: {feederConfig.MaxNumberOfPages}");
                    Console.WriteLine($"    - 支持双面扫描: {feederConfig.CanScanDuplex}");
                }
            }

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
                Console.WriteLine("  - AutoConfigured (自动配置)");
        }

        private static bool ValidateScanSource()
        {
            if (_scanner == null) return false;

            if (_options.Source == "flatbed")
                return _scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed);
            else if (_options.Source == "feeder")
                return _scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder);

            return false;
        }

        private static bool ValidateResolution()
        {
            return _options.Resolution >= 50 && _options.Resolution <= 1200;
        }

        private static async Task PerformScanAsync()
        {
            Console.WriteLine("\n=== 扫描配置 ===");
            Console.WriteLine($"扫描源: {_options.Source}");
            Console.WriteLine($"分辨率: {_options.Resolution} DPI");
            Console.WriteLine($"文件格式: {_options.Format}");
            Console.WriteLine($"颜色模式: {_options.ColorMode}");

            if (!string.IsNullOrEmpty(_options.OutputDirectory))
                Console.WriteLine($"输出目录: {_options.OutputDirectory}");

            Console.WriteLine();

            // 配置扫描仪
            ConfigureScanner();

            // 设置输出文件夹
            StorageFolder outputFolder;
            if (string.IsNullOrEmpty(_options.OutputDirectory))
            {
                outputFolder = ApplicationData.Current.LocalFolder;
            }
            else
            {
                try
                {
                    outputFolder = await StorageFolder.GetFolderFromPathAsync(_options.OutputDirectory);
                }
                catch
                {
                    Console.WriteLine($"无法访问输出目录 '{_options.OutputDirectory}'，使用默认目录。");
                    outputFolder = ApplicationData.Current.LocalFolder;
                }
            }

            // 开始扫描
            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                Console.WriteLine("正在扫描...");

                ImageScannerScanSource scanSource = _options.Source == "flatbed"
                    ? ImageScannerScanSource.Flatbed
                    : ImageScannerScanSource.Feeder;

                var progress = new Progress<uint>(pageCount =>
                {
                    Console.WriteLine($"已扫描 {pageCount} 页...");
                });

                var scanResult = await _scanner.ScanFilesToFolderAsync(
                    scanSource,
                    outputFolder
                ).AsTask(_cancellationTokenSource.Token, progress);

                // 显示结果
                if (scanResult.ScannedFiles.Count > 0)
                {
                    Console.WriteLine($"\n扫描完成! 共扫描 {scanResult.ScannedFiles.Count} 页。");

                    for (int i = 0; i < scanResult.ScannedFiles.Count; i++)
                    {
                        var file = scanResult.ScannedFiles[i];
                        Console.WriteLine($"文件 {i + 1}: {file.Name}");

                        try
                        {
                            var properties = await file.GetBasicPropertiesAsync();
                            Console.WriteLine($"  大小: {FormatFileSize(properties.Size)}");
                        }
                        catch { }
                    }
                }
                else
                {
                    Console.WriteLine("扫描完成，但未获得任何文件。");
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("扫描已取消。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"扫描失败: {ex.Message}");
                throw;
            }
        }

        private static void ConfigureScanner()
        {
            // 设置扫描源
            if (_options.Source == "flatbed")
            {
                ConfigureFlatbed();
            }
            else if (_options.Source == "feeder")
            {
                ConfigureFeeder();
            }
        }

        private static void ConfigureFlatbed()
        {
            var flatbedConfig = _scanner.FlatbedConfiguration;

            // 设置文件格式
            flatbedConfig.Format = ConvertFormat(_options.Format);

            // 设置颜色模式
            flatbedConfig.ColorMode = ConvertColorMode(_options.ColorMode);

            // 设置分辨率
            flatbedConfig.DesiredResolution = new ImageScannerResolution
            {
                DpiX = (uint)_options.Resolution,
                DpiY = (uint)_options.Resolution
            };
        }

        private static void ConfigureFeeder()
        {
            var feederConfig = _scanner.FeederConfiguration;

            // 设置文件格式
            feederConfig.Format = ConvertFormat(_options.Format);

            // 设置颜色模式
            feederConfig.ColorMode = ConvertColorMode(_options.ColorMode);

            // 设置分辨率
            feederConfig.DesiredResolution = new ImageScannerResolution
            {
                DpiX = (uint)_options.Resolution,
                DpiY = (uint)_options.Resolution
            };

            // 启用多页扫描
            feederConfig.MaxNumberOfPages = 0; // 0 表示扫描所有可用页
        }

        private static ImageScannerFormat ConvertFormat(string format)
        {
            return format.ToLower() switch
            {
                "pdf" => ImageScannerFormat.Pdf,
                "jpeg" or "jpg" => ImageScannerFormat.Jpeg,
                "png" => ImageScannerFormat.Png,
                "tiff" => ImageScannerFormat.Tiff,
                "bmp" => ImageScannerFormat.DeviceIndependentBitmap,
                _ => ImageScannerFormat.Pdf
            };
        }

        private static ImageScannerColorMode ConvertColorMode(string mode)
        {
            return mode.ToLower() switch
            {
                "lineart" => ImageScannerColorMode.Monochrome,
                "gray" => ImageScannerColorMode.Grayscale,
                "color" => ImageScannerColorMode.Color,
                _ => ImageScannerColorMode.Color
            };
        }

        private static string FormatFileSize(ulong bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }

    public class CommandLineOptions
    {
        public bool ListDevices { get; set; } = false;
        public string DeviceName { get; set; } = string.Empty;
        public int Resolution { get; set; } = 300;
        public string Source { get; set; } = string.Empty;
        public string OutputDirectory { get; set; } = string.Empty;
        public string Format { get; set; } = "pdf";
        public string ColorMode { get; set; } = "color";
    }
}