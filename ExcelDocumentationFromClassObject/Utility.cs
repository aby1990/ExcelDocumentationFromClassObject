using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

namespace ExcelDocumentationFromClassObject
{
    internal sealed class Utility
    {
        private DataTable dtDocument = new DataTable();

        private bool isInitital = true;

        public DataTable GetNameAndType<T>(T test, string name1 = "")
        {
            if (isInitital)
            {
                DataColumn[] dataColumns = new DataColumn[] { new DataColumn("Field", typeof(string)), new DataColumn("DataType", typeof(string)) };
                dtDocument.Columns.AddRange(dataColumns);
            }

            isInitital = false;
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(test))
            {
                string name = name1 + descriptor.Name;
                var type = descriptor.PropertyType;
                if (type.Name.ToLower().Contains("collection") && IsNeeded(descriptor.Name))
                {
                    try
                    {
                        string t = type.FullName.Substring(type.FullName.IndexOf("[[") + 2, type.FullName.IndexOf(",") - type.FullName.IndexOf("[[") - 2);
                        Type te = Type.GetType(t);
                        if (te == null)
                        {
                            t = type.FullName.Substring(type.FullName.IndexOf("[[") + 2, type.FullName.IndexOf("]]") - type.FullName.IndexOf("[[") - 2);
                            te = Type.GetType(t);
                        }
                        var op = Activator.CreateInstance(te);
                        GetNameAndType(op, $"{name}[].");
                        continue;
                    }
                    catch (Exception ex)
                    {
                    }
                }



                else if (!type.Module.Name.ToLower().Contains("corelib") && !type.Module.Name.ToLower().Contains("corlib")
                    && !type.Module.Name.ToLower().Contains("enum") && IsNeeded(descriptor.Name))
                {
                    try
                    {
                        var op = Activator.CreateInstance(type);
                        GetNameAndType(op, $"{name}.");
                        continue;
                    }
                    catch (Exception ex)
                    {
                    }
                }
                dtDocument.Rows.Add(name, type.FullName);
            }
            return dtDocument;
        }

        //To cut down size of excel, only chosen which are commonly used in any class, for 
        // small classes just return true to get complete object list 
        public bool IsNeeded(string str)
        {
            List<string> lst = new List<string>()
            {  
                "randomclass"
            };
            return true;// lst.Contains(str.ToLower()); 
        }
    }
}
