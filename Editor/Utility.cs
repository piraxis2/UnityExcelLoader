using UnityEngine;

namespace UnityExcelLoader.Editor
{
    public class Utility
    {
        public static string ConvertToUnityAssetPath(string fullPath)
        {
            string dataPath = Application.dataPath.Replace("/", "\\");
            if (fullPath.StartsWith(dataPath))
            {
                return "Assets" + fullPath.Substring(dataPath.Length);
            }

            Debug.LogError("The provided path is not within the Assets folder.");
            return null;
        } 
    }
}