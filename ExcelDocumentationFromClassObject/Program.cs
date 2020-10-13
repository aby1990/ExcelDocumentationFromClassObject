using ExcelDocumentationFromClassObject.SampleClasses;
using System;
using System.Data;

namespace ExcelDocumentationFromClassObject
{
    class Program
    {
        static void Main(string[] args)
        {
            Utility utility = new Utility();
            SampleClass sampleClass = new SampleClass(); //Change this to any type to get its fields 
            DataTable dtDocument = new DataTable();
            dtDocument = utility.GetNameAndType(sampleClass);
            //DataTable to excel functionality
            utility.ExportToExcel(dtDocument, $"Documentation_{DateTime.Now.ToString("MMddyyyyHHmm")}.xlsx", "DocumentCSharp");
        }
    }
}
