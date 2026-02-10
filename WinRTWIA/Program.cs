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

        static async Task Main(string[] args)
        {
            Console.WriteLine("=== WIA Scanner Command Line Tool ===");
            Console.WriteLine();

            try
            {
                // 查找扫描仪
                Console.Write("正在查找扫描仪...");
                var scanner = await FindScannerAsync();

                if (scanner == null)
                {
                    Console.WriteLine("未找到WIA兼容扫描仪!");
                    Console.WriteLine("请确保扫描仪已连接并开启。");
                    return;
                }

                Console.WriteLine($"找到扫描仪: {scanner.DeviceId}");
                _scanner = scanner;

                // 显示扫描仪信息
                await DisplayScannerInfoAsync();

                // 获取用户选择
                var options = await GetScanOptionsFromUserAsync();

                if (options == null)
                {
                    Console.WriteLine("扫描已取消。");
                    return;
                }

                // 执行扫描
                await PerformScanAsync(options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"内部错误: {ex.InnerException.Message}");
                }
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        private static async Task<ImageScanner> FindScannerAsync()
        {
            try
            {
                // 查找所有WIA扫描仪设备
                var deviceCollection = await DeviceInformation.FindAllAsync(ImageScanner.GetDeviceSelector());

                if (deviceCollection.Count == 0)
                {
                    return null;
                }

                // 返回第一个找到的扫描仪
                var deviceInfo = deviceCollection[0];
                return await ImageScanner.FromIdAsync(deviceInfo.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找扫描仪时出错: {ex.Message}");
                return null;
            }
        }

        private static async Task DisplayScannerInfoAsync()
        {
            Console.WriteLine("\n=== 扫描仪信息 ===");
            Console.WriteLine($"名称: {_scanner.DeviceId}");

            // 检查支持的扫描源
            Console.WriteLine("支持的扫描源:");
            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed))
            {
                Console.WriteLine("  - 平板扫描 (Flatbed)");
            }
            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder))
            {
                Console.WriteLine("  - 自动进纸器 (Feeder)");

                // 显示进纸器功能
                var feederConfig = _scanner.FeederConfiguration;
                Console.WriteLine(feederConfig.MaxNumberOfPages);
                Console.WriteLine($"  - 支持双面扫描: {feederConfig?.Duplex.ToString() ?? "未知"}");
                Console.WriteLine($"  - 支持多页扫描: {feederConfig?.MaxNumberOfPages > 1}");
            }
            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
            {
                Console.WriteLine("  - 自动配置 (Auto)");
            }

            Console.WriteLine();
        }

        private static async Task<ScanOptions> GetScanOptionsFromUserAsync()
        {
            var options = new ScanOptions();

            // 选择扫描源
            Console.WriteLine("选择扫描源:");
            Console.WriteLine("1. 平板扫描 (Flatbed)");
            Console.WriteLine("2. 自动进纸器 (Feeder)");
            Console.WriteLine("3. 自动配置 (Auto)");
            Console.Write("请选择 (1-3): ");

            var sourceChoice = Console.ReadLine();
            switch (sourceChoice)
            {
                case "1":
                    options.Source = ImageScannerScanSource.Flatbed;
                    break;
                case "2":
                    options.Source = ImageScannerScanSource.Feeder;
                    break;
                case "3":
                    options.Source = ImageScannerScanSource.AutoConfigured;
                    break;
                default:
                    Console.WriteLine("无效选择，使用默认设置。");
                    options.Source = ImageScannerScanSource.AutoConfigured;
                    break;
            }

            // 如果选择进纸器，询问多页扫描
            if (options.Source == ImageScannerScanSource.Feeder)
            {
                Console.Write("是否启用多页扫描? (y/n): ");
                var multiPageChoice = Console.ReadLine()?.ToLower();
                options.FeederMultiplePages = multiPageChoice == "y" || multiPageChoice == "yes";

                if (options.FeederMultiplePages)
                {
                    Console.Write("是否启用双面扫描? (y/n): ");
                    var duplexChoice = Console.ReadLine()?.ToLower();
                    options.FeederDuplex = duplexChoice == "y" || duplexChoice == "yes";
                }
            }

            // 选择文件格式
            Console.WriteLine("\n选择文件格式:");
            Console.WriteLine("1. PDF");
            Console.WriteLine("2. JPEG");
            Console.WriteLine("3. PNG");
            Console.WriteLine("4. TIFF");
            Console.WriteLine("5. BMP");
            Console.Write("请选择 (1-5): ");

            var formatChoice = Console.ReadLine();
            switch (formatChoice)
            {
                case "1":
                    options.Format = ImageScannerFormat.Pdf;
                    break;
                case "2":
                    options.Format = ImageScannerFormat.Jpeg;
                    break;
                case "3":
                    options.Format = ImageScannerFormat.Png;
                    break;
                case "4":
                    options.Format = ImageScannerFormat.Tiff;
                    break;
                case "5":
                    options.Format = ImageScannerFormat.DeviceIndependentBitmap;
                    break;
                default:
                    Console.WriteLine("无效选择，使用PDF格式。");
                    options.Format = ImageScannerFormat.Pdf;
                    break;
            }

            // 选择颜色模式
            Console.WriteLine("\n选择颜色模式:");
            Console.WriteLine("1. 彩色 (Color)");
            Console.WriteLine("2. 灰度 (Grayscale)");
            Console.WriteLine("3. 黑白 (Monochrome)");
            Console.WriteLine("4. 自动 (Auto)");
            Console.Write("请选择 (1-4): ");

            var colorChoice = Console.ReadLine();
            switch (colorChoice)
            {
                case "1":
                    options.ColorMode = ImageScannerColorMode.Color;
                    break;
                case "2":
                    options.ColorMode = ImageScannerColorMode.Grayscale;
                    break;
                case "3":
                    options.ColorMode = ImageScannerColorMode.Monochrome;
                    break;
                case "4":
                    options.ColorMode = ImageScannerColorMode.AutoColor;
                    break;
                default:
                    Console.WriteLine("无效选择，使用彩色模式。");
                    options.ColorMode = ImageScannerColorMode.Color;
                    break;
            }

            // 选择分辨率
            Console.WriteLine("\n选择分辨率 (DPI):");
            Console.WriteLine("1. 150 DPI (快速)");
            Console.WriteLine("2. 300 DPI (标准)");
            Console.WriteLine("3. 600 DPI (高质量)");
            Console.Write("请选择 (1-3): ");

            var dpiChoice = Console.ReadLine();
            switch (dpiChoice)
            {
                case "1":
                    options.Resolution = new ImageScannerResolution { DpiX = 150, DpiY = 150 };
                    break;
                case "2":
                    options.Resolution = new ImageScannerResolution { DpiX = 300, DpiY = 300 };
                    break;
                case "3":
                    options.Resolution = new ImageScannerResolution { DpiX = 600, DpiY = 600 };
                    break;
                default:
                    Console.WriteLine("无效选择，使用300 DPI。");
                    options.Resolution = new ImageScannerResolution { DpiX = 300, DpiY = 300 };
                    break;
            }

            // 选择保存位置
            Console.Write("\n输入保存文件夹路径 (留空使用当前文件夹): ");
            var folderPath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(folderPath))
            {
                options.TargetFolder = ApplicationData.Current.LocalFolder;
            }
            else
            {
                try
                {
                    options.TargetFolder = await StorageFolder.GetFolderFromPathAsync(folderPath);
                }
                catch
                {
                    Console.WriteLine("文件夹路径无效，使用当前文件夹。");
                    options.TargetFolder = ApplicationData.Current.LocalFolder;
                }
            }

            return options;
        }

        private static async Task PerformScanAsync(ScanOptions options)
        {
            Console.WriteLine("\n=== 开始扫描 ===");

            // 配置扫描仪
            ConfigureScanner(options);

            // 创建进度报告
            _cancellationTokenSource = new CancellationTokenSource();
            var progress = new Progress<uint>(pageCount =>
            {
                Console.WriteLine($"已扫描 {pageCount} 页...");
            });

            try
            {
                // 执行扫描
                Console.WriteLine("正在扫描，请稍候...");
                var scanResult = await _scanner.ScanFilesToFolderAsync(
                    options.Source,
                    options.TargetFolder
                ).AsTask(_cancellationTokenSource.Token, progress);
                
                // 检查结果
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

        private static void ConfigureScanner(ScanOptions options)
        {
            switch (options.Source)
            {
                case ImageScannerScanSource.Flatbed:
                    ConfigureFlatbed(options);
                    break;
                case ImageScannerScanSource.Feeder:
                    ConfigureFeeder(options);
                    break;
                case ImageScannerScanSource.AutoConfigured:
                    // 自动配置模式，使用默认设置
                    break;
            }
        }

        private static void ConfigureFlatbed(ScanOptions options)
        {
            var flatbedConfig = _scanner.FlatbedConfiguration;

            // 设置文件格式
            flatbedConfig.Format = options.Format;

            // 设置颜色模式
            flatbedConfig.ColorMode = options.ColorMode;

            // 设置分辨率
            flatbedConfig.DesiredResolution = options.Resolution;
        }

        private static void ConfigureFeeder(ScanOptions options)
        {
            var feederConfig = _scanner.FeederConfiguration;

            // 设置文件格式
            feederConfig.Format = options.Format;

            // 设置颜色模式
            feederConfig.ColorMode = options.ColorMode;

            // 设置分辨率
            feederConfig.DesiredResolution = options.Resolution;

            // 设置双面扫描
            if (options.FeederDuplex.HasValue)
            {
                feederConfig.Duplex = options.FeederDuplex.Value;
            }

            // 设置最大页数
            if (options.FeederMultiplePages)
            {
                if (feederConfig.Duplex)
                {
                    feederConfig.MaxNumberOfPages = 20; // 双面扫描最多20页
                }
                else
                {
                    feederConfig.MaxNumberOfPages = 10; // 单面扫描最多10页
                }
            }
            else
            {
                feederConfig.MaxNumberOfPages = 1; // 单页扫描
            }
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

    public class ScanOptions
    {
        public ImageScannerScanSource Source { get; set; }
        public ImageScannerFormat Format { get; set; }
        public ImageScannerColorMode ColorMode { get; set; }
        public ImageScannerResolution Resolution { get; set; }
        public bool FeederMultiplePages { get; set; }
        public bool? FeederDuplex { get; set; }
        public StorageFolder TargetFolder { get; set; }
    }
}