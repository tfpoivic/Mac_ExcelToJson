using System;
using System.IO;
using System.Text;
using Script.Data.Table;

namespace ExcelToJson {
    public class ExcelToJson {
        private const string JsonExt = ".json"; // json檔案副檔名
        private const string ExcelExt = ".xlsx"; // excel檔案副檔名

        public string DebugMessage { get; private set; } = string.Empty;

        private string FileListMessage { get; set; } = string.Empty;

        private readonly ExcelToJsonString _excelToJsonString;

        // Use this for initialization
        public ExcelToJson() {
            _excelToJsonString = new ExcelToJsonString();
        }

        // TODO:產生server和client檔案的區別
        public void TransferFilesFromExcelToJson(string excelDir, string jsonDir) {
            var clientDir = jsonDir + "//client";
            // var serverDir = jsonDir + "//server";
            if (!Directory.Exists(jsonDir)) // 如果資料夾不存在
            {
                Directory.CreateDirectory(jsonDir); // 建立目錄
                Directory.CreateDirectory(jsonDir + "//server"); // 建立目錄
                // Directory.CreateDirectory(jsonDir + "//client"); // 建立目錄
            }

            // if (!Directory.Exists(serverDir)) // 如果資料夾不存在
            // {
            //     Directory.CreateDirectory(serverDir); // 建立目錄
            // }

            if (!Directory.Exists(clientDir)) // 如果資料夾不存在
            {
                Directory.CreateDirectory(clientDir); // 建立目錄
            }

            var successFileCount = 0;

            var dataLoadTags = Enum.GetValues(typeof(EnumDataTables));
            var debugMsgBuilder = new StringBuilder();
            var tempDebugMsg = string.Empty;
            foreach (EnumDataTables dlt in dataLoadTags) {
                #region client

                var isSuccessGetAttr = GetAttribute<EnumClassValue>(dlt, out var dataConvertInfo);
                if (!isSuccessGetAttr) { continue; }

                var error = _excelToJsonString.ReadExcelFile(
                    excelDir,
                    dataConvertInfo,
                    NeedReadSite.CLIENT,
                    out var dataJsonString,
                    out tempDebugMsg
                );


                var excelFilePath = excelDir + Path.DirectorySeparatorChar + dataConvertInfo.FileName + ExcelExt;
                if (error == ReadExcelToJsonStringError.NONE) {
                    var jsonFilePath = clientDir + Path.DirectorySeparatorChar + dataConvertInfo.FileName + JsonExt;
                    WriteJsonStringToFile(dataJsonString, jsonFilePath);

                    debugMsgBuilder.AppendLine(string.Format("將 {0} 資料轉換成json成功", excelFilePath));
                    FileListMessage = string.Format("{0}{1}：O\n", FileListMessage, dataConvertInfo.FileName);
                    ++successFileCount;
                } else {
                    debugMsgBuilder.AppendLine(
                        string.Format("取得{0}內資料(型別為{1})失敗：失敗原因：{2}", excelFilePath, dataConvertInfo.ClassType, error)
                    );
                    FileListMessage = string.Format("{0}{1}：X\r\n", FileListMessage, dataConvertInfo.FileName);
                }

                #endregion
                //
                // #region server
                //
                // var isSuccessGetAttr2 = GetAttribute(dlt, out EnumClassValue dataConvertInfo2);
                // if (!isSuccessGetAttr2) { continue; }
                //
                // _excelToJsonString.ReadExcelFile(
                //     excelDir,
                //     dataConvertInfo2,
                //     NeedReadSite.SERVER,
                //     out var dataJsonString2,
                //     out tempDebugMsg
                // );
                //
                //
                // if (error == ReadExcelToJsonStringError.NONE) {
                //     var jsonFilePath2 =
                //         serverDir + Path.DirectorySeparatorChar + dataConvertInfo2.FileName + JsonExt;
                //     WriteJsonStringToFile(dataJsonString2, jsonFilePath2);
                //
                //     debugMsgBuilder.AppendLine(string.Format("將 {0} 資料轉換成json成功", excelFilePath));
                //     FileListMessage = string.Format("{0}{1}：O\n", FileListMessage, dataConvertInfo2.FileName);
                //     ++successFileCount;
                // } else {
                //     debugMsgBuilder.AppendLine(
                //         string.Format("取得{0}內資料(型別為{1})失敗：失敗原因：{2}", excelFilePath, dataConvertInfo.ClassType, error)
                //     );
                //     FileListMessage = string.Format("{0}{1}：X\r\n", FileListMessage, dataConvertInfo2.FileName);
                // }
                //
                // #endregion
            }

            debugMsgBuilder.AppendLine(
                string.Format("共轉換 {0}個檔案成功，{1}個檔案失敗", successFileCount, dataLoadTags.Length - successFileCount)
            );

            System.Diagnostics.Process.Start(clientDir);

            if (!string.IsNullOrEmpty(tempDebugMsg))
                debugMsgBuilder.AppendLine(string.Format("錯誤資訊\r\n{0}", tempDebugMsg));

            DebugMessage = debugMsgBuilder.ToString();
        }

        void WriteJsonStringToFile(string jsonString, string filePath) {
            using (StreamWriter sw = new StreamWriter(filePath)) {
                sw.Write(jsonString);
            }
        }

        private static bool GetAttribute<T>(Enum value, out T outAttr) where T : Attribute {
            outAttr = default(T);
            var curType = value.GetType();
            var curFieldInfo = curType.GetField(value.ToString());
            if (curFieldInfo == null) {
                return false;
            }

            if (!(curFieldInfo.GetCustomAttributes(typeof(T), false) is T[] curAttrs) || curAttrs.Length <= 0) {
                return false;
            }
            outAttr = curAttrs[0];
            return true;

        }
    }
}