using UnityEditor;
using UnityEngine;
using UnityExcelLoader.Runtime;

namespace UnityExcelLoader.Editor
{
    [CustomEditor(typeof(ScriptableObject))]
    public class UnityExcelLoaderEditor : UnityEditor.Editor
    {
        
        private Vector2 _scrollPos;

        public override void OnInspectorGUI()
        {
            var attributes = serializedObject.GetType().GetCustomAttributes(typeof(ExcelScriptableObject), false);
            if (attributes.Length == 0)
            {
                base.OnInspectorGUI();
                return;
            }
            
            serializedObject.Update();

            EditorGUILayout.LabelField("Excel Scriptable Object", EditorStyles.boldLabel);

            _scrollPos = EditorGUILayout.BeginScrollView(_scrollPos, GUILayout.Height(300));
            EditorGUILayout.PropertyField(serializedObject.FindProperty("data"), true);
            EditorGUILayout.EndScrollView();

            serializedObject.ApplyModifiedProperties();
        }
    }
}