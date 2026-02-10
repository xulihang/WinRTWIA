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
            Console.WriteLine("  -h, --help                Show this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  ScannerCLI -L");
            Console.WriteLine("  ScannerCLI -d \"HP Scanner\" -s flatbed -r 300 -f pdf -m Color -o C:\\Scans");
            Console.WriteLine("  ScannerCLI -s feeder -f jpeg -m Gray -r 150");
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

        private static async Task PerformScanAsync()
        {
            Console.WriteLine("\n=== Scan Configuration ===");
            Console.WriteLine($"Scan Source: {_options.Source}");
            Console.WriteLine($"Resolution: {_options.Resolution} DPI");
            Console.WriteLine($"File Format: {_options.Format}");
            Console.WriteLine($"Color Mode: {_options.ColorMode}");

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
            // Set scan source
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

            // Set file format
            flatbedConfig.Format = ConvertFormat(_options.Format);

            // Set color mode
            flatbedConfig.ColorMode = ConvertColorMode(_options.ColorMode);

            // Set resolution
            flatbedConfig.DesiredResolution = new ImageScannerResolution
            {
                DpiX = (uint)_options.Resolution,
                DpiY = (uint)_options.Resolution
            };
        }

        private static void ConfigureFeeder()
        {
            var feederConfig = _scanner.FeederConfiguration;

            // Set file format
            feederConfig.Format = ConvertFormat(_options.Format);

            // Set color mode
            feederConfig.ColorMode = ConvertColorMode(_options.ColorMode);

            // Set resolution
            feederConfig.DesiredResolution = new ImageScannerResolution
            {
                DpiX = (uint)_options.Resolution,
                DpiY = (uint)_options.Resolution
            };

            // Enable multi-page scanning
            feederConfig.MaxNumberOfPages = 0; // 0 means scan all available pages
            if (feederConfig.CanScanDuplex)
            {
                feederConfig.Duplex = true;
            }

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
    }
}