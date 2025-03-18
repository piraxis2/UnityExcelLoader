using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEngine;
using UnityEditor;

namespace UnityExcelLoader.Editor
{

    public struct ExcelDirectoryData
    {
        public string ExcelPath;
        public string ExcelName;
        public string SavePath;
    }
    public class UnityExcelLoaderMenuItems
    {
        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScript", false)]
        private static void ExcelToScript()
        {
            ExcelDirectoryData excelDirectoryData = GetExcelDirectoryData("Save ExcelScript");

            if (string.IsNullOrEmpty(excelDirectoryData.SavePath)) return;

            CreateEntity(excelDirectoryData);
            CreateScript(excelDirectoryData);
            AssetDatabase.Refresh();
        }

        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScript", true)]
        private static bool IsValidateCreateExcelScript() => IsExcel();
        
        
        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScriptableObject", false)]
        private static void ExcelToScriptableObject()
        {
            ExcelDirectoryData excelDirectoryData = GetExcelDirectoryData("Save ExcelScriptableObject");

            if (string.IsNullOrEmpty(excelDirectoryData.SavePath)) return;

            CreateScriptableObject(excelDirectoryData);
            AssetDatabase.Refresh();
        }

        [MenuItem("Assets/Create/UnityExcelLoader/ExcelScriptableObject", true)]
        private static bool IsValidateCreateExcelScriptableObject()
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            var excelPath = AssetDatabase.GetAssetPath(selectedAssets[0]);
            var excelName = Path.GetFileNameWithoutExtension(excelPath);
            var entityType = Type.GetType($"{excelName}_Entity, Assembly-CSharp");
            return IsExcel() && entityType != null;
        }


        private static void CreateScript(ExcelDirectoryData excelDirectoryData)
        {
            const string templatePath = "ScriptTemplate.txt";
            var filePath = Directory.GetFiles(Directory.GetCurrentDirectory(), templatePath, SearchOption.AllDirectories);
            if (filePath.Length == 0) throw new Exception("Script template not found.");

            var templateContent = File.ReadAllText(filePath[0]);
            var scriptPath = Path.Combine(excelDirectoryData.SavePath, $"{excelDirectoryData.ExcelName}.cs");

            var classContent = templateContent
                .Replace("{ClassName}", excelDirectoryData.ExcelName)
                .Replace("{Entity}", $"{excelDirectoryData.ExcelName}_Entity");

            if (File.Exists(scriptPath))
            {
                var existingContent = File.ReadAllText(scriptPath);
                var methodsStartIndex = existingContent.IndexOf("//MethodsStart", StringComparison.Ordinal) + "//MethodsStart".Length;
                var methodsEndIndex = existingContent.IndexOf("//MethodsEnd", StringComparison.Ordinal);
                var existingMethods = existingContent.Substring(methodsStartIndex, methodsEndIndex - methodsStartIndex).Trim();
                
                classContent = classContent.Replace("{Methods}", "\t" + existingMethods);
                
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

        private static void CreateEntity(ExcelDirectoryData excelDirectoryData)
        {
            using var stream = File.Open(excelDirectoryData.ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook book = Path.GetExtension(excelDirectoryData.ExcelPath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
            const string templatePath = "EntityTemplate.txt";
            var filePath = Directory.GetFiles(Directory.GetCurrentDirectory(), templatePath, SearchOption.AllDirectories);
            if (filePath.Length == 0) throw new Exception("Script template not found.");


            for (int i = 0; i < book.NumberOfSheets; i++)
            {
            	var templateContent = File.ReadAllText(filePath[0]);
                var sheet = book.GetSheetAt(i);
                var entityPath = Path.Combine(excelDirectoryData.SavePath, $"{excelDirectoryData.ExcelName}_Entity.cs");

                var headerRow = sheet.GetRow(0);
                var typeRow = sheet.GetRow(1);

                var fieldsBuilder = new StringBuilder();
                for (int j = 0; j < headerRow.LastCellNum; j++)
                {
                    var headerCell = headerRow.GetCell(j);
                    var typeCell = typeRow.GetCell(j);
                    
                    if (headerCell.StringCellValue.StartsWith("#")) continue;

                    if (headerCell.CellType != CellType.Blank && typeCell != null && typeCell.CellType != CellType.Blank)
                    {
                        fieldsBuilder.AppendLine($"\tpublic {typeCell.StringCellValue} {headerCell.StringCellValue};");
                    }
                }

                var classContent = templateContent
                    .Replace("{ClassName}", $"{excelDirectoryData.ExcelName}_Entity")
                    .Replace("{Fields}", fieldsBuilder.ToString());

                if (File.Exists(entityPath))
                {
                    var existingContent = File.ReadAllText(entityPath);
                    var methodsStartIndex = existingContent.IndexOf("//MethodsStart", StringComparison.Ordinal) + "//MethodsStart".Length;
                    var methodsEndIndex = existingContent.IndexOf("//MethodsEnd", StringComparison.Ordinal);
                    var existingMethods = existingContent.Substring(methodsStartIndex, methodsEndIndex - methodsStartIndex).Trim();

                    classContent = classContent.Replace("{Methods}", "\t" + existingMethods);

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

        private static void CreateScriptableObject(ExcelDirectoryData excelDirectoryData)
        {
            var unityPath = excelDirectoryData.SavePath[excelDirectoryData.SavePath.LastIndexOf("Assets", StringComparison.Ordinal)..];
            var scriptableObjectPath = unityPath + $"/{excelDirectoryData.ExcelName}.asset";
            var entityType = Type.GetType($"{excelDirectoryData.ExcelName}_Entity, Assembly-CSharp");

            if (entityType == null)
            {
                Debug.LogError($"Entity type {excelDirectoryData.ExcelName}_Entity not found.");
                return;
            }

            var scriptableObject = ScriptableObject.CreateInstance(excelDirectoryData.ExcelName);
            scriptableObject.hideFlags = HideFlags.NotEditable;
            var dataList = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(entityType));

            using var stream = File.Open(excelDirectoryData.ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook book = Path.GetExtension(excelDirectoryData.ExcelPath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);

            for (int i = 0; i < book.NumberOfSheets; i++)
            {
                var sheet = book.GetSheetAt(i);
                var headerRow = sheet.GetRow(0);
                
                for (int j = 2; j <= sheet.LastRowNum; j++)
                {
                    var row = sheet.GetRow(j);
                    var frontCell = row.GetCell(0);

                    if (frontCell.CellType == CellType.String && frontCell.StringCellValue.StartsWith("#")) continue;
                    
                    var entity = Activator.CreateInstance(entityType);

                    for (int k = 0; k < headerRow.LastCellNum; k++)
                    {
                        var headerCell = headerRow.GetCell(k);
                        
                        if (headerCell.StringCellValue.StartsWith("#")) continue;
                        
                        var dataCell = row.GetCell(k);

                        if (dataCell == null) continue;
                        
                        var fieldName = headerCell.StringCellValue;
                        var field = entityType.GetField(fieldName);
                        
                        if (field == null) continue;

                        if (string.IsNullOrEmpty(dataCell.ToString())) continue;
                        
                        try
                        {
                            var value = Convert.ChangeType(dataCell.ToString(), field.FieldType);
                            field.SetValue(entity, value);
                        }
                        catch (InvalidCastException)
                        {
                            throw new InvalidCastException($"Cannot convert '{dataCell}' to {field.FieldType.Name} for property '{fieldName}'.");
                        }
                        catch (FormatException)
                        {
                            throw new FormatException($"'{dataCell}' is not in a valid format for property '{fieldName}'.");
                        }
                        catch (OverflowException)
                        {
                            throw new OverflowException($"'{dataCell}' is not in a valid format for property '{fieldName}'.");
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"An error occurred while setting property '{fieldName}': {ex.Message}");
                        }
                    }

                    dataList.Add(entity);
                }
            }

            var fieldInfo = scriptableObject.GetType().GetField("Data");
            fieldInfo?.SetValue(scriptableObject, dataList);

            AssetDatabase.CreateAsset(scriptableObject, scriptableObjectPath);
            EditorUtility.SetDirty(scriptableObject);
        }

        private static bool IsExcel()
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            if (selectedAssets.Length != 1) return false;
            var path = AssetDatabase.GetAssetPath(selectedAssets[0]);
            return Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx";
        }

        private static ExcelDirectoryData GetExcelDirectoryData(string saveFolderTitle)
        {
            var selectedAssets = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.Assets);
            var excelPath = AssetDatabase.GetAssetPath(selectedAssets[0]);
            var excelName = Path.GetFileNameWithoutExtension(excelPath);
            var initialDirectory = Path.GetDirectoryName(excelPath);
            var savePath = EditorUtility.SaveFolderPanel(saveFolderTitle, initialDirectory, "");
            return new ExcelDirectoryData { ExcelPath = excelPath, ExcelName = excelName, SavePath = savePath };
        }
    }
}