using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Excel;
using UnityEditor;
using UnityEngine;

namespace ETCO
{
    public class ExcelToScriptableObject
    {
        [MenuItem("Config/Read Excel")]
        private static void ToExcelObject()
        {
            var fileList = Directory.GetFiles(Application.dataPath + "/hill_dash_hero/Editor/Excel/").ToList();
            fileList = fileList.FindAll(a => a.EndsWith("xlsx") && !a.Contains("~"));
            fileList.ForEach(a => Create(a));
        }

        private static void Create(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var data = reader.AsDataSet();

                    if (data == null)
                    {
                        Debug.LogError("read excel failed");
                        return;
                    }

                    foreach (DataTable table in data.Tables)
                    {
                        var tableName = table.TableName.Trim();
                        if (tableName.StartsWith("#")) continue;
                        CreateScript(table);
                    }

                    AssetDatabase.Refresh();
                    foreach (DataTable table in data.Tables)
                    {
                        var tableName = table.TableName.Trim();
                        if (tableName.StartsWith("#")) continue;
                        CreateObject(table);
                    }

                    AssetDatabase.Refresh();
                }
            }
        }

        private static void CreateScript(DataTable table)
        {
            var table_name = table.TableName.Trim().Replace("-", "_");
            var item_name = table_name + "Item";
            var str = new StringBuilder();
            var menu_str = $"[CreateAssetMenu(fileName = \"{table_name}\", menuName = \"Excel/{table_name}\", order = 0)]";

            str.AppendLine("//----------------------------------------------");
            str.AppendLine("//    Auto Generated. DO NOT edit manually!");
            str.AppendLine("//----------------------------------------------");
            str.AppendLine("");
            str.AppendLine("using System;");
            str.AppendLine("using UnityEngine;");
            str.AppendLine("using System.Collections.Generic;");
            str.AppendLine("using Sirenix.OdinInspector;");
            str.AppendLine("");
            str.AppendLine("namespace Game.Data");
            str.AppendLine("{");
            str.AppendLine($"    {menu_str}");
            str.AppendLine($"    public class {table_name} : ScriptableObject");
            str.AppendLine("    {");
            str.AppendLine("        [TableList(AlwaysExpanded = true)]");
            str.AppendLine($"        public List<{item_name}> _list = new List<{item_name}>();");
            str.AppendLine($"        public List<{item_name}> list {{");
            str.AppendLine("            set => _list = value;");
            str.AppendLine("            get => _list;");
            str.AppendLine("        }");
            str.AppendLine("");
            str.AppendLine("        #region 单例");
            str.AppendLine($"        private static {table_name} _self;");
            str.AppendLine($"        public static {table_name} self {{");
            str.AppendLine("            get {");
            str.AppendLine("                if (_self == null)");
            str.AppendLine("                {");
            str.AppendLine($"                    _self = Resources.Load<{table_name}>(\"Data/{table_name}\");");
            str.AppendLine("                }");
            str.AppendLine("                return _self;");
            str.AppendLine("            }");
            str.AppendLine("        }");
            str.AppendLine("        #endregion");
            str.AppendLine("");
            str.AppendLine("        [Serializable]");
            str.AppendLine($"        public class {item_name}");
            str.AppendLine("        {");

            for (var i = 0; i < table.Rows[0].ItemArray.Length; i++)
            {
                var des = table.Rows[0].ItemArray[i];
                var name = table.Rows[1].ItemArray[i].ToString().Replace("-", "_");
                var type = table.Rows[2].ItemArray[i];

                if (string.IsNullOrEmpty(name)) continue;

                str.AppendLine("");
                str.AppendLine("            /// <summary>");
                str.AppendLine($"            /// {des}");
                str.AppendLine("            /// <summary>");
                str.AppendLine($"            public {type} _{name};");
                str.AppendLine($"            public {type} {name} {{");
                str.AppendLine($"                set => _{name} = value;");
                str.AppendLine($"                get => _{name};");
                str.AppendLine("            }");
            }

            str.AppendLine("        }");
            str.AppendLine("    }");
            str.Append("}");

            var file = Application.dataPath + $"/hill_dash_hero/Scripts/Data/Excel/{table_name}.cs";

            if (!File.Exists(file)) File.Create(file);
            var bytes = new UTF8Encoding(false).GetBytes(str.ToString());
            File.WriteAllBytes(file, bytes);
        }

        private static void CreateObject(DataTable table)
        {
            var table_name = table.TableName.Trim().Replace("-", "_");
            var item_name = table_name + "Item";
            var type_table = Type.GetType($"Game.Data.{table_name}, Assembly-CSharp, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null");
            var type_item = Type.GetType($"Game.Data.{table_name}+{item_name}, Assembly-CSharp, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null");
            var sp = Resources.Load($"Data/{table_name}", type_table);

            var hang = table.Columns.Count;
            var lie = table.Rows.Count;
            // var list = new List<object>();
            var listV = typeof(List<>).MakeGenericType(type_item);
            var list = (IList)Activator.CreateInstance(listV);

            for (var y = 3; y < lie; y++)
            {
                if (type_item != null)
                {
                    var item = Activator.CreateInstance(type_item);

                    for (var x = 0; x < hang; x++)
                    {
                        var name = table.Rows[1].ItemArray[x].ToString().Replace("-", "_");
                        if (string.IsNullOrEmpty(name)) break;

                        var value = table.Rows[y].ItemArray[x];
                        Debug.Log($"{x}_{y}_{name}_{value}");

                        var dtype = item.GetType().GetProperty(name).PropertyType;
                        if (dtype == typeof(int))
                        {
                            if (value is DBNull)
                                item.GetType().GetProperty(name).SetValue(item, 0);
                            else
                                item.GetType().GetProperty(name).SetValue(item, Convert.ToInt32(value));
                        }
                        else if (dtype == typeof(string))
                        {
                            if (value is DBNull)
                                item.GetType().GetProperty(name).SetValue(item, "");
                            else
                                item.GetType().GetProperty(name).SetValue(item, ((string)value).Trim());
                        }
                        else if (dtype == typeof(float))
                        {
                            if (value is DBNull)
                                item.GetType().GetProperty(name).SetValue(item, 0f);
                            else
                                item.GetType().GetProperty(name).SetValue(item, Convert.ToSingle(value));
                        }
                        else
                            item.GetType().GetProperty(name).SetValue(item, Convert.ToSingle(value));
                    }

                    list.Add(item);
                }
            }

            var ti = type_table.GetProperty("list");
            ti.SetValue(sp, list);

            EditorUtility.SetDirty(sp);
        }
    }
}