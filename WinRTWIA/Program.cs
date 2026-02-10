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
using Windows.Foundation;

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
                // Parse command line arguments
                if (!ParseArguments(args))
                {
                    ShowHelp();
                    return;
                }

                // If listing devices is specified
                if (_options.ListDevices)
                {
                    await ListScannersAsync();
                    return;
                }

                // Find scanner
                ImageScanner scanner;
                if (!string.IsNullOrEmpty(_options.DeviceName))
                {
                    scanner = await FindScannerByNameAsync(_options.DeviceName);
                    if (scanner == null)
                    {
                        Console.WriteLine($"Scanner device with name '{_options.DeviceName}' not found!");
                        return;
                    }
                }
                else
                {
                    scanner = await FindFirstScannerAsync();
                    if (scanner == null)
                    {
                        Console.WriteLine("No WIA-compatible scanner found!");
                        Console.WriteLine("Please ensure the scanner is connected and powered on.");
                        return;
                    }
                }

                _scanner = scanner;

                // Validate if scan source is supported
                if (!ValidateScanSource())
                {
                    Console.WriteLine($"Scanner does not support '{_options.Source}' scan source!");
                    await DisplayScannerInfoAsync();
                    return;
                }

                // Validate resolution
                if (!ValidateResolution())
                {
                    Console.WriteLine("Resolution must be between 50-1200 DPI!");
                    return;
                }

                // Validate scan area
                if (!ValidateScanArea())
                {
                    Console.WriteLine("Invalid scan area parameters!");
                    return;
                }

                // Validate contrast and brightness values
                if (!ValidateImageEnhancements())
                {
                    Console.WriteLine("Contrast and brightness values must be between -100 and 100!");
                    return;
                }

                // Perform scan
                await PerformScanAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner error: {ex.InnerException.Message}");
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
                            Console.WriteLine("Error: -d parameter requires device name");
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
                                Console.WriteLine("Error: Resolution must be a number");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -r parameter requires resolution value");
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
                                Console.WriteLine("Error: Scan source must be 'flatbed' or 'feeder'");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -s parameter requires scan source");
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
                            Console.WriteLine("Error: -o parameter requires output directory");
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
                                Console.WriteLine("Error: Unsupported file format");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -f parameter requires file format");
                            return false;
                        }
                        break;

                    case "-m":
                    case "--mode":
                        if (i + 1 < args.Length)
                        {
                            _options.ColorMode = args[++i].ToLower();
                            if (!new[] { "lineart", "grayscale", "color" }.Contains(_options.ColorMode))
                            {
                                Console.WriteLine("Error: Unsupported color mode");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -m parameter requires color mode");
                            return false;
                        }
                        break;

                    case "-c":
                    case "--contrast":
                        if (i + 1 < args.Length)
                        {
                            if (int.TryParse(args[++i], out int contrast))
                            {
                                _options.Contrast = contrast;
                            }
                            else
                            {
                                Console.WriteLine("Error: Contrast must be a number between -100 and 100");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -c parameter requires contrast value");
                            return false;
                        }
                        break;

                    case "-b":
                    case "--brightness":
                        if (i + 1 < args.Length)
                        {
                            if (int.TryParse(args[++i], out int brightness))
                            {
                                _options.Brightness = brightness;
                            }
                            else
                            {
                                Console.WriteLine("Error: Brightness must be a number between -100 and 100");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -b parameter requires brightness value");
                            return false;
                        }
                        break;

                    // 扫描区域参数
                    case "-l":
                    case "--left":
                        if (i + 1 < args.Length)
                        {
                            if (float.TryParse(args[++i], out float left))
                            {
                                _options.ScanAreaLeft = left;
                                _options.ScanAreaSpecified = true;
                            }
                            else
                            {
                                Console.WriteLine("Error: Left position must be a number (millimeters)");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -l parameter requires left position value");
                            return false;
                        }
                        break;

                    case "-t":
                    case "--top":
                        if (i + 1 < args.Length)
                        {
                            if (float.TryParse(args[++i], out float top))
                            {
                                _options.ScanAreaTop = top;
                                _options.ScanAreaSpecified = true;
                            }
                            else
                            {
                                Console.WriteLine("Error: Top position must be a number (millimeters)");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -t parameter requires top position value");
                            return false;
                        }
                        break;

                    case "-x":
                    case "--width":
                        if (i + 1 < args.Length)
                        {
                            if (float.TryParse(args[++i], out float width))
                            {
                                _options.ScanAreaWidth = width;
                                _options.ScanAreaSpecified = true;
                            }
                            else
                            {
                                Console.WriteLine("Error: Width must be a number (millimeters)");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -x parameter requires width value");
                            return false;
                        }
                        break;

                    case "-y":
                    case "--height":
                        if (i + 1 < args.Length)
                        {
                            if (float.TryParse(args[++i], out float height))
                            {
                                _options.ScanAreaHeight = height;
                                _options.ScanAreaSpecified = true;
                            }
                            else
                            {
                                Console.WriteLine("Error: Height must be a number (millimeters)");
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Error: -y parameter requires height value");
                            return false;
                        }
                        break;

                    case "-h":
                    case "--help":
                        ShowHelp();
                        return false;

                    default:
                        Console.WriteLine($"Unknown parameter: {args[i]}");
                        ShowHelp();
                        return false;
                }
            }

            // Validate required parameters (if not listing devices)
            if (!_options.ListDevices)
            {
                if (string.IsNullOrEmpty(_options.Source))
                {
                    Console.WriteLine("Error: Must specify scan source (-s)");
                    return false;
                }
            }

            return true;
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Scanner Command Line Tool");
            Console.WriteLine("Usage: ScannerCLI [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -L, --list                List all available scanner devices");
            Console.WriteLine("  -d, --device NAME         Specify scanner device to use by name");
            Console.WriteLine("  -r, --resolution DPI      Specify scan resolution (default: 300)");
            Console.WriteLine("  -s, --source SOURCE       Select scan source: flatbed or feeder");
            Console.WriteLine("  -o, --output DIR          Specify output directory (default: current directory)");
            Console.WriteLine("  -f, --format FORMAT       Specify file format: pdf, jpeg, png, tiff, bmp");
            Console.WriteLine("  -m, --mode MODE           Select color mode: Lineart, Gray, Color");
            Console.WriteLine("  -c, --contrast VALUE      Adjust contrast (-100 to 100, default: 0)");
            Console.WriteLine("  -b, --brightness VALUE    Adjust brightness (-100 to 100, default: 0)");
            Console.WriteLine("  -l, --left MM             Left position of scan area in millimeters");
            Console.WriteLine("  -t, --top MM              Top position of scan area in millimeters");
            Console.WriteLine("  -x, --width MM            Width of scan area in millimeters");
            Console.WriteLine("  -y, --height MM           Height of scan area in millimeters");
            Console.WriteLine("  -h, --help                Show this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  ScannerCLI -L");
            Console.WriteLine("  ScannerCLI -d \"HP Scanner\" -s flatbed -r 300 -f pdf -m Color -o C:\\Scans");
            Console.WriteLine("  ScannerCLI -s feeder -f jpeg -m Gray -r 150");
            Console.WriteLine("  ScannerCLI -s flatbed -r 300 -f jpeg -l 10 -t 10 -x 200 -y 280");
            Console.WriteLine("  ScannerCLI -s flatbed -l 20 -t 20 -x 100 -y 150");
            Console.WriteLine("  ScannerCLI -s flatbed -r 300 -c 20 -b 10  # Increase contrast and brightness");
            Console.WriteLine("  ScannerCLI -s flatbed -r 300 -c -10 -b -5 # Decrease contrast and brightness");
        }

        private static async Task ListScannersAsync()
        {
            Console.WriteLine("Searching for scanner devices...\n");

            try
            {
                var deviceCollection = await DeviceInformation.FindAllAsync(ImageScanner.GetDeviceSelector());

                if (deviceCollection.Count == 0)
                {
                    Console.WriteLine("No scanner devices found.");
                    return;
                }

                Console.WriteLine($"Found {deviceCollection.Count} device(s):\n");

                for (int i = 0; i < deviceCollection.Count; i++)
                {
                    var device = deviceCollection[i];
                    Console.WriteLine($"[{i + 1}] {device.Name}");
                    Console.WriteLine($"  ID: {device.Id}");

                    try
                    {
                        var scanner = await ImageScanner.FromIdAsync(device.Id);
                        Console.WriteLine("  Supported scan sources:");

                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed))
                            Console.WriteLine("    - Flatbed");
                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder))
                            Console.WriteLine("    - Feeder");
                        if (scanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
                            Console.WriteLine("    - Auto");
                    }
                    catch
                    {
                        Console.WriteLine("  Unable to retrieve device details");
                    }

                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while searching for devices: {ex.Message}");
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
                Console.WriteLine($"Error finding scanner: {ex.Message}");
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
                Console.WriteLine($"Error finding scanner: {ex.Message}");
                return null;
            }
        }

        private static async Task DisplayScannerInfoAsync()
        {
            if (_scanner == null) return;

            Console.WriteLine("\n=== Scanner Information ===");

            var deviceInfo = await DeviceInformation.CreateFromIdAsync(_scanner.DeviceId);
            Console.WriteLine($"Device Name: {deviceInfo.Name}");
            Console.WriteLine($"Device ID: {_scanner.DeviceId}");
            Console.WriteLine("Supported scan sources:");

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Flatbed))
                Console.WriteLine("  - Flatbed");

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.Feeder))
            {
                Console.WriteLine("  - Feeder");
                var feederConfig = _scanner.FeederConfiguration;
                if (feederConfig != null)
                {
                    Console.WriteLine($"    - Max pages: {feederConfig.MaxNumberOfPages}");
                    Console.WriteLine($"    - Supports duplex scanning: {feederConfig.CanScanDuplex}");
                }
            }

            if (_scanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
                Console.WriteLine("  - AutoConfigured");
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

        private static bool ValidateScanArea()
        {
            // 如果没有指定扫描区域，则使用默认（整个扫描区域）
            if (!_options.ScanAreaSpecified)
                return true;

            // 检查是否所有四个参数都已指定
            if (_options.ScanAreaLeft < 0 || _options.ScanAreaTop < 0)
            {
                Console.WriteLine("Left and top positions must be >= 0");
                return false;
            }

            if (_options.ScanAreaWidth <= 0 || _options.ScanAreaHeight <= 0)
            {
                Console.WriteLine("Width and height must be > 0");
                return false;
            }

            return true;
        }

        private static bool ValidateImageEnhancements()
        {
            // 验证对比度和亮度值在有效范围内 (-100 到 100)
            if (_options.Contrast < -1000 || _options.Contrast > 1000)
            {
                Console.WriteLine("Contrast must be between -1000 and 1000");
                return false;
            }

            if (_options.Brightness < -1000 || _options.Brightness > 1000)
            {
                Console.WriteLine("Brightness must be between -1000 and 1000");
                return false;
            }

            return true;
        }

        private static async Task PerformScanAsync()
        {
            Console.WriteLine("\n=== Scan Configuration ===");
            Console.WriteLine($"Scan Source: {_options.Source}");
            Console.WriteLine($"Resolution: {_options.Resolution} DPI");
            Console.WriteLine($"File Format: {_options.Format}");
            Console.WriteLine($"Color Mode: {_options.ColorMode}");

            if (_options.Contrast != 0)
                Console.WriteLine($"Contrast: {_options.Contrast}");

            if (_options.Brightness != 0)
                Console.WriteLine($"Brightness: {_options.Brightness}");

            if (_options.ScanAreaSpecified)
            {
                Console.WriteLine($"Scan Area: Left={_options.ScanAreaLeft}mm, Top={_options.ScanAreaTop}mm, " +
                                  $"Width={_options.ScanAreaWidth}mm, Height={_options.ScanAreaHeight}mm");
            }
            else
            {
                Console.WriteLine($"Scan Area: Full scan bed");
            }

            if (!string.IsNullOrEmpty(_options.OutputDirectory))
                Console.WriteLine($"Output Directory: {_options.OutputDirectory}");

            Console.WriteLine();

            // Configure scanner
            ConfigureScanner();

            // Set output folder
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
                    Console.WriteLine($"Cannot access output directory '{_options.OutputDirectory}', using default directory.");
                    outputFolder = ApplicationData.Current.LocalFolder;
                }
            }

            // Start scanning
            _cancellationTokenSource = new CancellationTokenSource();

            try
            {
                Console.WriteLine("Scanning...");

                ImageScannerScanSource scanSource = _options.Source == "flatbed"
                    ? ImageScannerScanSource.Flatbed
                    : ImageScannerScanSource.Feeder;
                
                var progress = new Progress<uint>(pageCount =>
                {
                    Console.WriteLine($"Scanned {pageCount} page(s)...");
                });

                var scanResult = await _scanner.ScanFilesToFolderAsync(
                    scanSource,
                    outputFolder
                ).AsTask(_cancellationTokenSource.Token, progress);

                // Display results
                if (scanResult.ScannedFiles.Count > 0)
                {
                    Console.WriteLine($"\nScan completed! Total {scanResult.ScannedFiles.Count} page(s) scanned.");

                    for (int i = 0; i < scanResult.ScannedFiles.Count; i++)
                    {
                        var file = scanResult.ScannedFiles[i];
                        Console.WriteLine($"File {i + 1}: {file.Name}");

                        try
                        {
                            var properties = await file.GetBasicPropertiesAsync();
                            Console.WriteLine($"  Size: {FormatFileSize(properties.Size)}");
                        }
                        catch { }
                    }
                }
                else
                {
                    Console.WriteLine("Scan completed, but no files were obtained.");
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Scan cancelled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Scan failed: {ex.Message}");
                throw;
            }
        }

        private static void ConfigureScanner()
        {
            Console.WriteLine($"Configuring");
            // Set scan source
            if (_options.Source == "flatbed")
            {
                ConfigureFlatbed();
            }
            else if (_options.Source == "feeder")
            {
                ConfigureFeeder();
            }
            Console.WriteLine($"Configure Done");
        }

        private static void ConfigureFlatbed()
        {
            var flatbedConfig = _scanner.FlatbedConfiguration;
            Console.WriteLine("configure format");
            // Set file format
            try
            {
                flatbedConfig.Format = ConvertFormat(_options.Format);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set format: {ex.Message}");
                flatbedConfig.Format = ImageScannerFormat.DeviceIndependentBitmap;
            }
            Console.WriteLine("configure color");
            // Set color mode
            try
            {
                flatbedConfig.ColorMode = ConvertColorMode(_options.ColorMode);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set color mode: {ex.Message}");
            }
            Console.WriteLine("configure resolution");
            // Set resolution
            try
            {
                flatbedConfig.DesiredResolution = new ImageScannerResolution
                {
                    DpiX = (uint)_options.Resolution,
                    DpiY = (uint)_options.Resolution
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set dpi: {ex.Message}");
            }

            // Set contrast if specified
            if (_options.Contrast != 0)
            {
                try
                {
                    // 将对比度值从 -100..100 映射到实际的对比度范围
                    // 注意：实际范围可能因扫描仪而异，这里使用相对调整
                    flatbedConfig.Contrast = NormalizeEnhancementValue(_options.Contrast);
                    Console.WriteLine($"Applied contrast: {flatbedConfig.Contrast}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not set contrast: {ex.Message}");
                }
            }

            // Set brightness if specified
            if (_options.Brightness != 0)
            {
                try
                {
                    // 将亮度值从 -100..100 映射到实际的亮度范围
                    flatbedConfig.Brightness = NormalizeEnhancementValue(_options.Brightness);
                    Console.WriteLine($"Applied brightness: {flatbedConfig.Brightness}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not set brightness: {ex.Message}");
                }
            }

            // Set scan area if specified
            if (_options.ScanAreaSpecified)
            {
                // 将毫米转换为英寸（1英寸 = 25.4毫米）
                float dpiX = _options.Resolution;
                float dpiY = _options.Resolution;

                // 计算像素值：毫米 * DPI / 25.4
                float leftInch = _options.ScanAreaLeft / 25.4f;
                float topInch = _options.ScanAreaTop / 25.4f;
                float widthInch = _options.ScanAreaWidth / 25.4f;
                float heightInch = _options.ScanAreaHeight / 25.4f;

                // 设置扫描区域（以像素为单位）
                flatbedConfig.SelectedScanRegion = new Rect(
                    leftInch,
                    topInch,
                    widthInch,
                    heightInch
                );
                Console.WriteLine($"Scan region in inches: Left={leftInch}, Top={topInch}, " +
                                $"Width={widthInch}, Height={heightInch}");
            }
        }

        private static void ConfigureFeeder()
        {
            var feederConfig = _scanner.FeederConfiguration;
            Console.WriteLine("configure format");
            // Set file format
            try
            {
                feederConfig.Format = ConvertFormat(_options.Format);
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Warning: Could not set format: {ex.Message}");
                feederConfig.Format = ImageScannerFormat.DeviceIndependentBitmap;
            }
            Console.WriteLine("configure color");
            // Set color mode
            try
            {
                feederConfig.ColorMode = ConvertColorMode(_options.ColorMode);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set color mode: {ex.Message}");
            }
            Console.WriteLine("configure resolution");
            // Set resolution
            try {
                feederConfig.DesiredResolution = new ImageScannerResolution
                {
                    DpiX = (uint)_options.Resolution,
                    DpiY = (uint)_options.Resolution
                };
            }
            catch (Exception ex) {
                Console.WriteLine($"Warning: Could not set dpi: {ex.Message}");
            }
            

            // Set contrast if specified
            if (_options.Contrast != 0)
            {
                try
                {
                    feederConfig.Contrast = NormalizeEnhancementValue(_options.Contrast);
                    Console.WriteLine($"Applied contrast: {feederConfig.Contrast}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not set contrast: {ex.Message}");
                }
            }

            // Set brightness if specified
            if (_options.Brightness != 0)
            {
                try
                {
                    feederConfig.Brightness = NormalizeEnhancementValue(_options.Brightness);
                    Console.WriteLine($"Applied brightness: {feederConfig.Brightness}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not set brightness: {ex.Message}");
                }
            }

            // Set scan area if specified
            if (_options.ScanAreaSpecified)
            {
                // 将毫米转换为英寸（1英寸 = 25.4毫米）
                float dpiX = _options.Resolution;
                float dpiY = _options.Resolution;

                // 计算像素值：毫米 * DPI / 25.4
                float leftInch = _options.ScanAreaLeft / 25.4f;
                float topInch = _options.ScanAreaTop / 25.4f;
                float widthInch = _options.ScanAreaWidth / 25.4f;
                float heightInch = _options.ScanAreaHeight / 25.4f;

                uint leftPixels = (uint)(leftInch * dpiX);
                uint topPixels = (uint)(topInch * dpiY);
                uint widthPixels = (uint)(widthInch * dpiX);
                uint heightPixels = (uint)(heightInch * dpiY);

                // 设置扫描区域（以像素为单位）
                feederConfig.SelectedScanRegion = new Rect(
                    leftPixels,
                    topPixels,
                    widthPixels,
                    heightPixels
                );
                Console.WriteLine($"Scan region in pixels: Left={leftPixels}, Top={topPixels}, " +
                                $"Width={widthPixels}, Height={heightPixels}");
            }

            // Enable multi-page scanning
            feederConfig.MaxNumberOfPages = 0; // 0 means scan all available pages
            if (feederConfig.CanScanDuplex)
            {
                feederConfig.Duplex = true;
            }
        }

        private static int NormalizeEnhancementValue(int value)
        {
            return value;
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
                _ => ImageScannerFormat.DeviceIndependentBitmap
            };
        }

        private static ImageScannerColorMode ConvertColorMode(string mode)
        {
            return mode.ToLower() switch
            {
                "lineart" => ImageScannerColorMode.Monochrome,
                "grayscale" => ImageScannerColorMode.Grayscale,
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

        // 图像增强参数
        public int Contrast { get; set; } = 0;
        public int Brightness { get; set; } = 0;

        // 扫描区域参数
        public float ScanAreaLeft { get; set; } = 0;
        public float ScanAreaTop { get; set; } = 0;
        public float ScanAreaWidth { get; set; } = 0;
        public float ScanAreaHeight { get; set; } = 0;
        public bool ScanAreaSpecified { get; set; } = false;
    }
}