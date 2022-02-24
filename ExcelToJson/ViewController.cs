using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using AppKit;
using Foundation;
using MobileCoreServices;
using Xamarin.Essentials;

namespace ExcelToJson {
    public partial class ViewController : NSViewController {

        private string _filePath = "";
        
        public ViewController(IntPtr handle) : base(handle) { }

        public override void ViewDidLoad() {
            base.ViewDidLoad();
            // Do any additional setup after loading the view.
        }

        public override NSObject RepresentedObject {
            get { return base.RepresentedObject; }
            set {
                base.RepresentedObject = value;
                // Update the view, if already loaded.
            }
        }

        static FilePickerFileType PlatformXlsxFileType() =>
            new FilePickerFileType(
                new Dictionary<DevicePlatform, IEnumerable<string>> {{DevicePlatform.macOS, new string[] {".xlsx"}}}
            );

        partial void selectInputPath(AppKit.NSButtonCell sender) {
            Console.WriteLine("select");
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
            convertExcelToJson();
        }

        private async void convertExcelToJson() {
            ExcelToJson etj = new ExcelToJson();
            await Task.Run(
                () => {
                    etj.TransferFilesFromExcelToJson(_filePath, _filePath);
                }
            );
            logTextView.Value = etj.DebugMessage;
        }
    }
}