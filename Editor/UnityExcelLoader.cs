using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

namespace UnityExcelLoader.Editor
{
    public class UnityExcelLoader : AssetPostprocessor
    {
        private static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
        {
            foreach (var asset in importedAssets)
            {
                if (Path.GetExtension(asset) == ".xlsx" || Path.GetExtension(asset) == ".xls")
                {
                    LoadExcel(asset);
                }
            }
        }

        private static void LoadExcel(string asset)
        {
            var excelName = Path.GetFileNameWithoutExtension(asset);
            var scriptableObjectPath = $"Assets/{excelName}.asset";
            var entityType = Type.GetType($"UnityExcelLoader.{excelName}_Entity, Assembly-CSharp");

            if (entityType == null)
            {
                Debug.LogError($"Entity type {excelName}_Entity not found.");
                return;
            }

            var scriptableObject = ScriptableObject.CreateInstance(excelName);
            scriptableObject.hideFlags = HideFlags.NotEditable;
            var dataList = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(entityType));

            using var stream = File.Open(asset, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook book = Path.GetExtension(asset) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);

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
    }
}