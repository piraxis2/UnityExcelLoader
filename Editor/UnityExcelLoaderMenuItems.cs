using System.IO;
using System;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEngine;
using UnityEditor;

namespace UnityExcelLoader.Editor
{
    public class UnityExcelLoaderMenuItems
    {
        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScript", false)]
        private static void ExcelToScriptableObject()
        {
            var savePath = EditorUtility.SaveFolderPanel("Save ExcelAssetScript", Application.dataPath, "");
            if (string.IsNullOrEmpty(savePath)) return;

            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            var excelPath = AssetDatabase.GetAssetPath(selectedAssets[0]);
            var excelName = Path.GetFileNameWithoutExtension(excelPath);

            CreateEntity(savePath, excelPath, excelName);
            CreateScript(savePath, excelName);
            AssetDatabase.Refresh();
        }

        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScript", true)]
        private static bool IsValidateCreateExcelScript() => IsExcel();

        private static void CreateScript(string savePath, string excelName)
        {
            const string templatePath = "ScriptTemplate.txt";
            var filePath = Directory.GetFiles(Directory.GetCurrentDirectory(), templatePath, SearchOption.AllDirectories);
            if (filePath.Length == 0) throw new Exception("Script template not found.");

            var templateContent = File.ReadAllText(filePath[0]);
            var scriptPath = Path.Combine(savePath, $"{excelName}.cs");

            var classContent = templateContent
                .Replace("{ClassName}", excelName)
                .Replace("{Entity}", $"{excelName}_Entity");

            if (File.Exists(scriptPath))
            {
                var existingContent = File.ReadAllText(scriptPath);
                var methodsStartIndex = existingContent.IndexOf("//MethodsStart", StringComparison.Ordinal) + "//MethodsStart".Length;
                var methodsEndIndex = existingContent.IndexOf("//MethodsEnd", StringComparison.Ordinal);
                var existingMethods = existingContent.Substring(methodsStartIndex, methodsEndIndex - methodsStartIndex).Trim();
                
                classContent = classContent.Replace("{Methods}", "\t\t" + existingMethods);
                
                var customUsing =  existingContent.IndexOf("//CustomUsingStart", StringComparison.Ordinal) + "//CustomUsingStart".Length;
                var customUsingEnd = existingContent.IndexOf("//CustomUsingEnd", StringComparison.Ordinal);
                var existingCustomUsing = existingContent.Substring(customUsing, customUsingEnd - customUsing).Trim();
                
                classContent = classContent.Replace("{CustomUsing}", existingCustomUsing);
            }
            else
            {
                classContent = classContent.Replace("{Methods}", "");
                classContent = classContent.Replace("{CustomUsing}", "");
            }

            File.WriteAllText(scriptPath, classContent);
        }

        private static void CreateEntity(string savePath, string excelPath, string excelName)
        {
            using var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook book = Path.GetExtension(excelPath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
            const string templatePath = "EntityTemplate.txt";
            var filePath = Directory.GetFiles(Directory.GetCurrentDirectory(), templatePath, SearchOption.AllDirectories);
            if (filePath.Length == 0) throw new Exception("Script template not found.");


            for (int i = 0; i < book.NumberOfSheets; i++)
            {
            	var templateContent = File.ReadAllText(filePath[0]);
                var sheet = book.GetSheetAt(i);
                var entityPath = Path.Combine(savePath, $"{excelName}_Entity.cs");

                var headerRow = sheet.GetRow(0);
                var typeRow = sheet.GetRow(1);

                var fieldsBuilder = new StringBuilder();
                for (int j = 0; j < headerRow.LastCellNum; j++)
                {
                    var headerCell = headerRow.GetCell(j);
                    var typeCell = typeRow.GetCell(j);

                    if (headerCell != null && headerCell.CellType != CellType.Blank &&
                        typeCell != null && typeCell.CellType != CellType.Blank)
                    {
                        fieldsBuilder.AppendLine($"\t\tpublic {typeCell.StringCellValue} {headerCell.StringCellValue};");
                    }
                }

                var classContent = templateContent
                    .Replace("{ClassName}", $"{excelName}_Entity")
                    .Replace("{Fields}", fieldsBuilder.ToString());

                if (File.Exists(entityPath))
                {
                    var existingContent = File.ReadAllText(entityPath);
                    var methodsStartIndex = existingContent.IndexOf("//MethodsStart", StringComparison.Ordinal) + "//MethodsStart".Length;
                    var methodsEndIndex = existingContent.IndexOf("//MethodsEnd", StringComparison.Ordinal);
                    var existingMethods = existingContent.Substring(methodsStartIndex, methodsEndIndex - methodsStartIndex).Trim();

                    classContent = classContent.Replace("{Methods}", "\t\t" + existingMethods);

                    var customUsing = existingContent.IndexOf("//CustomUsingStart", StringComparison.Ordinal) + "//CustomUsingStart".Length;
                    var customUsingEnd = existingContent.IndexOf("//CustomUsingEnd", StringComparison.Ordinal);
                    var existingCustomUsing = existingContent.Substring(customUsing, customUsingEnd - customUsing).Trim();

                    classContent = classContent.Replace("{CustomUsing}", existingCustomUsing);
                }
                else
                {
                    classContent = classContent.Replace("{Methods}", "");
                    classContent = classContent.Replace("{CustomUsing}", "");
                }

                File.WriteAllText(entityPath, classContent);
            }
        }

        private static bool IsExcel()
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            if (selectedAssets.Length != 1) return false;
            var path = AssetDatabase.GetAssetPath(selectedAssets[0]);
            return Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx";
        }
    }
}