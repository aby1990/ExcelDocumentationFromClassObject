using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
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

        public void ExportToExcel(DataTable DataTable, string fileName, string sheetName, bool isSameSheet = false)
        {
            try
            {
                string parser = GetParentDirectory() + @"\Response\" + fileName;
                bool isFileExists = false;
                int ColumnsCount;
                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");
                // load excel, and create a new workbook 

                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb;
                Microsoft.Office.Interop.Excel._Worksheet Worksheet;

                if (File.Exists(parser))
                {
                    isFileExists = true;
                    wb = Excel.Workbooks.Open(parser);
                    Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)(isSameSheet ? Excel.ActiveSheet : wb.Sheets.Add());
                }
                else
                {
                    wb = Excel.Workbooks.Add();
                    Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)Excel.ActiveSheet;
                }
                Worksheet.Name = sheetName;
                
                object[] Header = new object[ColumnsCount];

                int lastUsedRow = 0;
                int index = 2;
                if (File.Exists(parser) && isSameSheet)
                {
                    index = 1;
                    lastUsedRow = Worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                  Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                  false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;                }

                if (lastUsedRow == 0)
                {
                    for (int i = 0; i < ColumnsCount; i++)
                        Header[i] = DataTable.Columns[i].ColumnName;

                    Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                    HeaderRange.Value = Header;
                    //HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray); 
                    HeaderRange.Font.Bold = true;

                }

                // Microsoft.Office.Interop.Excel.Range range = Worksheet.UsedRange;
                // DataCells 

                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j, i] = DataTable.Rows[j][i];

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[lastUsedRow + index, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + lastUsedRow + index - 1, ColumnsCount])).Value = Cells;
                // check fielpath 
                if (parser != null && parser != "")
                {
                    try
                    {
                        if (isFileExists)
                            wb.Save();
                        else
                            Worksheet.SaveAs(parser);
                        Excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"+ ex.Message);
                    }
                }
                else    // no filepath is given 
                {
                    Excel.Quit();
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }



        public string GetParentDirectory()
        {
            string ParafilePath = AppDomain.CurrentDomain.BaseDirectory;
            DirectoryInfo ParaparentDir = Directory.GetParent(ParafilePath);
            return ParaparentDir.Parent.FullName;
        }
    }
}
