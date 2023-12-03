using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using LevelsCreationExcel.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LevelsCreationExcel.RevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class OpenExcelWindowCommand : IExternalCommand
    {
        #region Properties
        public static Document Doc { get; set; }
        public static UIDocument uIDocument { get; set; }
        #endregion
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uIDocument = commandData.Application.ActiveUIDocument;

            Doc = uIDocument.Document;
            try
            {
                ExcelWindow mainWindow = new ExcelWindow();

                mainWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;

                return Result.Failed;

            }
        }
    }
}
