using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
namespace ExcelToJson
{	
	public class ExcelToJson
	{
		readonly string JSON_EXT = ".json";        // json檔案副檔名
		readonly string EXCEL_EXT = ".xlsx";       // excel檔案副檔名

		string _debugMessage = string.Empty;
		public string DebugMessage
		{
			get { return _debugMessage; }
		}
		string _fileListMessage = string.Empty;
		public string FileListMessage
		{
			get { return _fileListMessage; }
		}
		bool _currentlyTransfering = false; // 是否正在轉換中
		public bool CurrentlyTransfering
		{
			get { return _currentlyTransfering; }
		}
		ExcelToJsonString excelToJsonString;

		// Use this for initialization
		public ExcelToJson()
		{
			excelToJsonString = new ExcelToJsonString();
		}

		// TODO:產生server和client檔案的區別
		public void TransferFilesFromExcelToJson(string excelDir, string jsonDir)
		{
            string clientDir = jsonDir + "//client";
            string serverDir = jsonDir + "//server";
			_currentlyTransfering = true;
			if (!Directory.Exists(jsonDir)) // 如果資料夾不存在
			{
				Directory.CreateDirectory(jsonDir); // 建立目錄
                Directory.CreateDirectory(jsonDir+"//server"); // 建立目錄
                Directory.CreateDirectory(jsonDir + "//client"); // 建立目錄
			}
            if (!Directory.Exists(serverDir)) // 如果資料夾不存在
            {
                Directory.CreateDirectory(serverDir); // 建立目錄
            }
            if (!Directory.Exists(clientDir)) // 如果資料夾不存在
            {
                Directory.CreateDirectory(clientDir); // 建立目錄
            }
			int successFileCount = 0;

			Array dataLoadTags = Enum.GetValues(typeof(EnumDataTables));
			StringBuilder debugMSGBuilder = new StringBuilder();
			string tempDebugMsg = string.Empty;
			foreach (EnumDataTables dlt in dataLoadTags)
            {
                #region client
                string dataJsonString;
				EnumClassValue dataConvertInfo;
				bool isSuccessGetAttr = GetAttribute<EnumClassValue>(dlt, out dataConvertInfo);
				if (!isSuccessGetAttr) { continue; }

				ReadExcelToJsonStringError error = excelToJsonString.ReadExcelFile(excelDir, dataConvertInfo, NeedReadSite.CLIENT, out dataJsonString, out tempDebugMsg);
				

				string excelFilePath = excelDir + Path.DirectorySeparatorChar + dataConvertInfo.FileName + EXCEL_EXT;
				if (error == ReadExcelToJsonStringError.NONE)
				{
                    string jsonFilePath = clientDir + Path.DirectorySeparatorChar + dataConvertInfo.FileName + JSON_EXT;
					WriteJsonStringToFile(dataJsonString, jsonFilePath);

					debugMSGBuilder.AppendLine(string.Format("將 {0} 資料轉換成json成功",  excelFilePath));
					_fileListMessage = string.Format("{0}{1}：O\n", _fileListMessage, dataConvertInfo.FileName);
					++successFileCount;
				}
				else
				{
					debugMSGBuilder.AppendLine(string.Format("取得{0}內資料(型別為{1})失敗：失敗原因：{2}", excelFilePath, dataConvertInfo.ClassType, error));
					_fileListMessage = string.Format("{0}{1}：X\r\n", _fileListMessage, dataConvertInfo.FileName);
                }
                #endregion

                #region server
                string dataJsonString2;
                EnumClassValue dataConvertInfo2;
                bool isSuccessGetAttr2 = GetAttribute<EnumClassValue>(dlt, out dataConvertInfo2);
                if (!isSuccessGetAttr2) { continue; }

                ReadExcelToJsonStringError error2 = excelToJsonString.ReadExcelFile(excelDir, dataConvertInfo2, NeedReadSite.SERVER, out dataJsonString2, out tempDebugMsg);


                string excelFilePath2 = excelDir + Path.DirectorySeparatorChar + dataConvertInfo2.FileName + EXCEL_EXT;
                if (error == ReadExcelToJsonStringError.NONE)
                {
                    string jsonFilePath2 = serverDir + Path.DirectorySeparatorChar + dataConvertInfo2.FileName + JSON_EXT;
                    WriteJsonStringToFile(dataJsonString2, jsonFilePath2);

                    debugMSGBuilder.AppendLine(string.Format("將 {0} 資料轉換成json成功", excelFilePath));
                    _fileListMessage = string.Format("{0}{1}：O\n", _fileListMessage, dataConvertInfo2.FileName);
                    ++successFileCount;
                }
                else
                {
                    debugMSGBuilder.AppendLine(string.Format("取得{0}內資料(型別為{1})失敗：失敗原因：{2}", excelFilePath, dataConvertInfo.ClassType, error));
                    _fileListMessage = string.Format("{0}{1}：X\r\n", _fileListMessage, dataConvertInfo2.FileName);
                }
                #endregion
            }
			debugMSGBuilder.AppendLine(string.Format("共轉換 {0}個檔案成功，{1}個檔案失敗", successFileCount, dataLoadTags.Length - successFileCount));

            System.Diagnostics.Process.Start(clientDir);

			if(!string.IsNullOrEmpty(tempDebugMsg))
				debugMSGBuilder.AppendLine(string.Format("錯誤資訊\r\n{0}", tempDebugMsg));

			_debugMessage = debugMSGBuilder.ToString();
			_currentlyTransfering = false;


		}

		void WriteJsonStringToFile(string jsonString, string filePath)
		{
			using (StreamWriter sw = new StreamWriter(filePath))
			{
				sw.Write(jsonString);
			}
		}

		public static bool GetAttribute<T>(System.Enum value, out T outAttr) where T : System.Attribute
		{
			outAttr = default(T);
			System.Type curType = value.GetType();
			System.Reflection.FieldInfo curFieldInfo = curType.GetField(value.ToString());
			if (curFieldInfo != null)
			{
				T[] curAttrs = curFieldInfo.GetCustomAttributes(typeof(T), false) as T[];
				if (curAttrs != null && curAttrs.Length > 0)
				{
					outAttr = curAttrs[0];
					return true;
				}
			}
			return false;
		}
	}
}
