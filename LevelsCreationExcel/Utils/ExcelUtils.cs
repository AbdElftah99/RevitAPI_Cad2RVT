using Autodesk.Revit.DB;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LevelsCreationExcel.Utils
{
    public class ExcelUtils
    {
        public static string OpenExcelAndGetFilePath()
        {
            Microsoft.Win32.OpenFileDialog openDialog = new Microsoft.Win32.OpenFileDialog();
            openDialog.Title = "Select an Excel File";
            openDialog.Filter = "Excel Files (.xlsx)|*.xlsx";
            openDialog.Multiselect = false;

            openDialog.RestoreDirectory = true;
            if (openDialog.ShowDialog() == true)
            {
                string fileName = openDialog.FileName;
                // Show the file path in the textbox
                return fileName;
            }
            else
            {
                return string.Empty;
            }
        }

        public static List<string> RetriveSheetsNames(string filepath)
        {
            List<string> SheetsNames = new List<string>();
            ExcelPackage package = new ExcelPackage(filepath);
            using (package)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets;
                foreach ( var row in sheet )
                {
                    SheetsNames.Add(row.Name);
                }
                return SheetsNames;
            }
        }
    }
}
