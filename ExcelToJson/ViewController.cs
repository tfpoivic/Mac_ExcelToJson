using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AppKit;
using Xamarin.Essentials;

namespace ExcelToJson {
    public partial class ViewController : NSViewController {

        private string _filePath = "";
        
        public ViewController(IntPtr handle) : base(handle) { }

        private static FilePickerFileType PlatformXlsxFileType() =>
            new FilePickerFileType(
                new Dictionary<DevicePlatform, IEnumerable<string>> {{DevicePlatform.macOS, new[] {".xlsx"}}}
            );

        partial void selectInputPath(NSButtonCell sender) {
            try {
                var inputPath = FilePicker.PickAsync(
                        new PickOptions {FileTypes = PlatformXlsxFileType()}
                    )
                    .GetAwaiter().GetResult().FullPath;
                var splitIndex = inputPath.LastIndexOfAny(new[] {'/'});
                var path = inputPath.Substring(0, splitIndex);
                inputPathTextField.StringValue = path;
                _filePath = path;
            } catch (Exception e) {
                Console.WriteLine(e);
            }
        }

        partial void startConvertExcelToJson(NSButton sender) {
            if (_filePath.Length == 0) {
                return;
            }
            logTextView.Value = "產檔中…";
            ConvertExcelToJson();
        }

        private async void ConvertExcelToJson() {
            var excelToJson = new ExcelToJson();
            await Task.Run(
                () => {
                    excelToJson.TransferFilesFromExcelToJson(_filePath, _filePath);
                }
            );
            logTextView.Value = excelToJson.DebugMessage;
        }
    }
}