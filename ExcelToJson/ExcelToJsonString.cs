using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Script.Data.Table;

namespace ExcelToJson {
    /// <summary>
    /// 負責將table內容(一連串的string)轉換成json格式。
    /// </summary>
    public class ExcelToJsonString {
        private readonly Dictionary<Type, string> _baseTypeString = new Dictionary<Type, string>() {
            {typeof(byte), "BYTE"},
            {typeof(ushort), "USHORT"},
            {typeof(uint), "UINT"},
            {typeof(ulong), "ULONG"},
            {typeof(string), "STRING"},
            {typeof(bool), "BOOL"},
            {typeof(float), "FLOAT"},
            {typeof(int), "INT"},
        };

        private string _debugMessage = string.Empty;
        private ExcelToTable _excelToTable;

        private List<int> _dontNeedColumnIndexes; // table不需要的欄位Index

        public ExcelToJsonString() {
            _excelToTable = new ExcelToTable();
        }

        ~ExcelToJsonString() {
            _excelToTable = null;
        }

        #region Excel讀取

        /// <summary>
        /// 處理讀取完excel檔案（不論成功還失敗）需要做的事情
        /// </summary>
        /// <param name="objDataForExcel">從excel取得的物件資料</param>
        /// <param name="jsonString">對應objDataForExcel要轉換成的json字串</param>
        /// <param name="debugString">偵錯字串</param>
        private void ReadExcelFileEnd(object objDataForExcel, out string jsonString, out string debugString) {
            jsonString = (objDataForExcel == null) ? string.Empty : SerializeObject(objDataForExcel);
            debugString = _debugMessage;
            _excelToTable.Close();
        }

        /// <summary>
        /// 讀取excel檔案，不論是否有錯誤，回傳前都會關閉檔案
        /// </summary>
        /// <param name="directoryPath">資料夾路徑</param>
        /// <param name="dataConvertInfo">資料轉換的資訊</param>
        /// <param name="needReadSite">轉出資料是哪方（Server/Client）需要</param>
        /// <param name="jsonString">對應輸出的json字串</param>
        /// <param name="debugString">偵錯字串</param>
        /// <returns>可能有的錯誤訊息</returns>
        public ReadExcelToJsonStringError ReadExcelFile(
            string directoryPath,
            EnumClassValue dataConvertInfo,
            NeedReadSite needReadSite,
            out string jsonString,
            out string debugString
        ) {
            if (string.IsNullOrEmpty(dataConvertInfo.FileName) || dataConvertInfo.ClassType == null) {
                ReadExcelFileEnd(null, out jsonString, out debugString);
                return ReadExcelToJsonStringError.ENUM_ATTRIBUTE_ERROR;
            }

            var readExcelError = _excelToTable.OpenExcelFile(
                directoryPath,
                dataConvertInfo.FileName
            );
            if (readExcelError != ReadExcelToJsonStringError.NONE) {
                ReadExcelFileEnd(null, out jsonString, out debugString);
                return readExcelError;
            }

            readExcelError = _excelToTable.CheckAndReadTableHeader(
                needReadSite,
                out var titleNames,
                out var allType,
                out _dontNeedColumnIndexes
            );
            if (readExcelError != ReadExcelToJsonStringError.NONE) {
                ReadExcelFileEnd(null, out jsonString, out debugString);
                return readExcelError;
            }

            #region 確認各欄位和要被寫入的物件欄位Type有對應

            var checkObject = Activator.CreateInstance(dataConvertInfo.ClassType);
            var isConform = CheckObjectTypeCorrect(
                dataConvertInfo.ClassType,
                checkObject,
                titleNames,
                allType
            );
            if (!isConform) {
                _debugMessage =
                    $"{_debugMessage} {dataConvertInfo.FileName} 轉換失敗：表格與資料結構({dataConvertInfo.ClassType})內容不符\r\n";
                ReadExcelFileEnd(null, out jsonString, out debugString);
                return ReadExcelToJsonStringError.TABLE_TYPE_IS_NOT_CONFORM;
            }

            #endregion

            #region 抓取資料

            var allData = new List<object>();

            var hasEor = false;

            var tableRowData = _excelToTable.GetNextRow();
            while (tableRowData != null) // 還有資料
            {
                if (ExcelToTable.CheckEndOfTable(tableRowData)) {
                    // 有結尾符號，正常結束
                    hasEor = true;
                    break;
                }

                if (!CheckEmptyRow(tableRowData) && !string.IsNullOrEmpty(tableRowData[0])) //空的欄位略過處理
                {
                    object obj = Activator.CreateInstance(dataConvertInfo.ClassType);
                    List<string>.Enumerator rowDataEnumerator = tableRowData.GetEnumerator();
                    ReadExcelToJsonStringError error = GetObjectTypeDataFromExcel(
                        dataConvertInfo.ClassType,
                        ref obj,
                        ref titleNames,
                        ref rowDataEnumerator
                    );
                    if (error != ReadExcelToJsonStringError.NONE) {
                        ReadExcelFileEnd(null, out jsonString, out debugString);
                        return error;
                    }

                    allData.Add(obj);
                    // 取得下一行資料
                }

                tableRowData = _excelToTable.GetNextRow();
            }

            if (!hasEor) {
                ReadExcelFileEnd(null, out jsonString, out debugString);
                return ReadExcelToJsonStringError.CANT_FIND_END_OF_ROW_TOKEN;
            }

            #endregion

            ReadExcelFileEnd(allData, out jsonString, out debugString);
            return ReadExcelToJsonStringError.NONE;
        }

        #endregion

        #region 確認型別對應

        /// <summary>
        /// 確認excel表格內定義的Type是否和給予的物件資料結構有對應
        /// </summary>
        /// <param name="checkType">給予的type定義</param>
        /// <param name="checkObject">對應的物件(不可為null)</param>
        /// <param name="titleNames">table中欄位對應的標題</param>
        /// <param name="tableTypeEnumerator">table內文</param>
        /// <returns>是否有對應</returns>
        private bool CheckObjectTypeCorrect(
            IReflect checkType,
            object checkObject,
            IReadOnlyList<string> titleNames,
            IReadOnlyList<string> tableTypeEnumerator
        ) {
            var allFieldInfo = checkType.GetFields(BindingFlags.Public | BindingFlags.Instance);
            var titleToFieldInfo =
                allFieldInfo.ToDictionary(
                    fieldInfo => fieldInfo.GetCustomAttribute<TitleName>()?.GetTitle() ?? fieldInfo.Name
                );
            for (var index = 0; index < titleNames.Count; index++) {
                var titleName = titleNames[index];
                var tableType = tableTypeEnumerator[index];
                if (!titleToFieldInfo.TryGetValue(titleName, out var fieldInfo)) {
                    continue;
                }

                if (_dontNeedColumnIndexes.Exists(x => x == index)) {
                    continue;
                }

                var curType = fieldInfo.FieldType;
                // 避免curObj為Null，否則可能會讓後面method取不到資訊(string為例外狀況)
                var curObj = fieldInfo.GetValue(checkObject);
                if (curType != typeof(string) && curObj == null) { Activator.CreateInstance(curType); }

                if (curType.IsArray || curType.IsClass && curType != typeof(string)) {
                    continue;
                }

                var isConform = CheckBaseTypeCorrect(curType, tableType);
                
                if (isConform) {
                    continue;
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// 確認excel表格內定義的Type是否和給予的基本資料結構有對應
        /// </summary>
        /// <param name="checkType">給予的type定義</param>
        /// <param name="tableType">table內文</param>
        /// <returns>是否有對應</returns>
        private bool CheckBaseTypeCorrect(Type checkType, string tableType) {

            // 由於可能有nullable型態，取得對應的非nullable型態再比較
            var isNullableType = checkType.IsGenericType && checkType.GetGenericTypeDefinition() == typeof(Nullable<>);
            var realType = (isNullableType) ? Nullable.GetUnderlyingType(checkType) : checkType;

            if (realType != null && _baseTypeString.TryGetValue(realType, out var compareStr)) // 是基本四型態之一
            {
                if (tableType != null
                    && tableType.ToUpper().Equals(compareStr.ToUpper())) {
                    return true;
                }

                Console.WriteLine($"base error : Type = {realType} excelType = {tableType}");
                _debugMessage =
                    $"{_debugMessage} base error : Type = {realType} excelType = {tableType}";
                return false;
            }

            Console.WriteLine($"not base error : Type = {realType}");
            _debugMessage = $"{_debugMessage} not base error : Type = {realType}";
            return false;
        }

        #endregion

        /// <summary>
        /// 確定是否為空行（整行資料都是沒文字或null）
        /// </summary>
        /// <param name="rowData">整行的資料</param>
        /// <returns>是否為空行</returns>
        private bool CheckEmptyRow(List<string> rowData) {
            if (rowData == null || rowData.Count == 0) { return true; }

            return rowData.All(string.IsNullOrEmpty);
        }

        #region 取得對應型別資料

        /// <summary>
        /// 從excel檔案中取得物件型別的資料
        /// </summary>
        /// <param name="type">物件型別的type</param>
        /// <param name="retObj">存放取得的資料</param>
        /// <param name="titleNames">table中欄位對應的標題</param>
        /// <param name="rowDataEnumerator">由excel來的row Data</param>
        /// <returns>可能的錯誤</returns>
        private ReadExcelToJsonStringError GetObjectTypeDataFromExcel(
            IReflect type,
            ref object retObj,
            ref List<string> titleNames,
            ref List<string>.Enumerator rowDataEnumerator
        ) {
            var error = ReadExcelToJsonStringError.NONE;
            var isNull = true;
            var allFieldInfo = type.GetFields(BindingFlags.Public | BindingFlags.Instance);
            var titleToFieldInfo =
                allFieldInfo.ToDictionary(
                    fieldInfo => fieldInfo.GetCustomAttribute<TitleName>()?.GetTitle() ?? fieldInfo.Name
                );
            for (var index = 0; index < titleNames.Count; index++) {
                var titleName = titleNames[index];
                if (!titleToFieldInfo.TryGetValue(titleName, out var fieldInfo)) {
                    rowDataEnumerator.MoveNext();
                    continue;
                }

                if (_dontNeedColumnIndexes.Exists(x => x == index)) {
                    fieldInfo.SetValue(retObj, null);
                } else {
                    var curFieldType = fieldInfo.FieldType;
                    // 避免curFieldObj為Null，否則可能會讓後面method取不到資訊(string為例外狀況)
                    var curFieldObj = fieldInfo.GetValue(retObj);
                    if (curFieldType != typeof(string) && curFieldObj == null) {
                        curFieldObj = Activator.CreateInstance(curFieldType);
                    }

                    if (curFieldType.IsArray || curFieldType.IsClass && curFieldType != typeof(string)) {
                        error = ReadExcelToJsonStringError.GET_BASE_TYPE_ERROR;
                    } else {
                        error = GetBaseTypeDataFromExcel(
                            curFieldType,
                            out curFieldObj,
                            ref rowDataEnumerator
                        );
                    }

                    if (error != ReadExcelToJsonStringError.NONE) {
                        retObj = null;
                        return error;
                    }

                    if (curFieldObj != null) { isNull = false; }

                    fieldInfo.SetValue(retObj, curFieldObj);
                    if (isNull) { retObj = null; }
                }
            }

            return error;
        }

        /// <summary>
        /// 從excel檔案中取得基本型別的資料
        /// </summary>
        /// <param name="type">基本型別的type</param>
        /// <param name="retObj">存放取得的資料</param>
        /// <param name="rowDataEnumerator">由excel來的row Data</param>
        /// <returns>可能的錯誤</returns>
        static ReadExcelToJsonStringError GetBaseTypeDataFromExcel(
            Type type,
            out object retObj,
            ref List<string>.Enumerator rowDataEnumerator
        ) {
            if (!rowDataEnumerator.MoveNext()) {
                retObj = null;
                return ReadExcelToJsonStringError.NONE;
            }

            //消空白，應該不會有人刻意要填空白，應該都是不小心的
            var isNull = rowDataEnumerator.Current == null || string.IsNullOrEmpty(rowDataEnumerator.Current.Trim());

            if (type == typeof(string)) {
                retObj = (isNull) ? null : rowDataEnumerator.Current;
                return ReadExcelToJsonStringError.NONE;
            } else {
                var isNullableType = type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
                if (isNull) // 資料是空的
                {
                    //實值型別為空資料的補0
                    if (type.IsValueType) {
                        retObj = Convert.ChangeType(0, type);
                        return ReadExcelToJsonStringError.NONE;
                    }

                    retObj = null;
                    return (isNullableType)
                        ? ReadExcelToJsonStringError.NONE
                        : ReadExcelToJsonStringError.DATA_CANT_BE_NULL; // type可為null，則資料為空沒問題
                }

                object[] para = {rowDataEnumerator.Current};
                var transferType = new[] {typeof(string)};
                var realType = (isNullableType) ? Nullable.GetUnderlyingType(type) : type;
                try {
                    retObj = realType?.GetMethod("Parse", transferType)?.Invoke(null, para);
                    return ReadExcelToJsonStringError.NONE;
                } catch (Exception e) {
                    Console.WriteLine($"Type: {transferType[0]}, Content:{para[0]}");
                    Console.WriteLine("取得基本型別時出錯 {2} \n{0}\n{1}", e.Message, e.StackTrace, type);
                    retObj = null;
                    return ReadExcelToJsonStringError.GET_BASE_TYPE_ERROR;
                }
            }
        }

        #endregion

        /// <summary>
        /// 將物件內的資料序列化為字串
        /// </summary>
        /// <param name="ob">要序列化的物件</param>
        /// <returns>序列化後的字串</returns>
        private static string SerializeObject(object ob) {
            var settings = new Newtonsoft.Json.JsonSerializerSettings {
                ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore, CheckAdditionalContent = false
            };
            return Newtonsoft.Json.JsonConvert.SerializeObject(ob, settings);
        }
    }
}