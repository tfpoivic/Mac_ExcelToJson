// WARNING
//
// This file has been generated automatically by Rider IDE
//   to store outlets and actions made in Xcode.
// If it is removed, they will be lost.
// Manual changes to this file may not be handled correctly.
//
using Foundation;
using System.CodeDom.Compiler;

namespace ExcelToJson
{
	[Register ("ViewController")]
	partial class ViewController
	{
		[Outlet]
		AppKit.NSTextField inputPathTextField { get; set; }

		[Outlet]
		AppKit.NSTextView logTextView { get; set; }

		[Action ("selectInputPath:")]
		partial void selectInputPath (AppKit.NSButtonCell sender);

		[Action ("startConvertExcelToJson:")]
		partial void startConvertExcelToJson (AppKit.NSButton sender);

		void ReleaseDesignerOutlets ()
		{
			if (inputPathTextField != null) {
				inputPathTextField.Dispose ();
				inputPathTextField = null;
			}

			if (logTextView != null) {
				logTextView.Dispose ();
				logTextView = null;
			}

		}
	}
}
