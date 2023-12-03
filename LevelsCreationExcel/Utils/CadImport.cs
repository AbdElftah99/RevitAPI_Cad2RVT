using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LevelsCreationExcel.Utils
{
    public class CadImport
    {
        public static IList<ElementId> GetImportedCad(Document doc)
        {
            IList<ElementId> cadimports;
            cadimports = (IList<ElementId>)new FilteredElementCollector(doc).OfClass(typeof(ImportInstance))
                .WhereElementIsNotElementType().ToElementIds();

            return cadimports;
        }
    }
}
