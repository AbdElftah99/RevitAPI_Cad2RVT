using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using LevelsCreationExcel.Commands;
using LevelsCreationExcel.RevitCommands;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Media3D;

namespace LevelsCreationExcel.ViewModel
{
    public class FootingCreateViewModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void onPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Fields&Properties
        Document Doc = OpenFootingWindowCommand.Doc;

        private IList<PolyLine> plines = new List<PolyLine>();
        //private IList<Arc> arcs = new List<Arc>();

        public IList<string> layersname { get; set; } = new List<string>();
        public IList<string> levelsname { get; set; } = new List<string>();
        public IList<string> type { get; set; } = new List<string>();

        private string _selectedlayer;
        public string Selectedlayer
        {
            get { return _selectedlayer; }
            set
            {

                _selectedlayer = value;

                onPropertyChanged(nameof(Selectedlayer));
            }
        }

        private string _selectedlevel;
        public string Selectedlevel
        {
            get { return _selectedlevel; }
            set
            {

                _selectedlevel = value;

                onPropertyChanged(nameof(Selectedlevel));
            }
        }

        private string _selectedtype;
        public string selectedtype
        {
            get { return _selectedtype; }
            set
            {

                _selectedtype = value;

                onPropertyChanged(nameof(selectedtype));
            }
        }


        private string _depth;
        public string Depth
        {
            get { return _depth; }
            set
            {

                _depth = value;

                onPropertyChanged();
            }
        }
        #endregion

        #region Commands
        public MyCommand CreateFootingCommand { get; set; }
        #endregion

        #region Constructor
        public FootingCreateViewModel()
        {
            CadLoad();
            CreateFootingCommand = new MyCommand(ExecuteCreateFoundationCommand , CanExecuteCreateFoundationCommand);
        }
        #endregion

        #region Methods
        public void CadLoad()
        {
            IList<ElementId> cadimports = (IList<ElementId>)new FilteredElementCollector(Doc).OfClass(typeof(ImportInstance))
                .WhereElementIsNotElementType().ToElementIds();

            IList<Level> levels = new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            foreach (var level in levels)
            {
                levelsname.Add(level.Name);
            }

            FilteredElementCollector collector = new FilteredElementCollector(Doc);
            ICollection<Element> elements = collector.OfCategory(BuiltInCategory.OST_StructuralFoundation).ToElements();

            foreach (Element element in elements)
            {
                if (element is ElementType elementType)
                {
                    type.Add(element.Name);
                }
            }

            if (cadimports.Count > 0)
            {
                ImportInstance imp = Doc.GetElement(cadimports.First()) as ImportInstance;
                GeometryElement geoel = imp.get_Geometry(new Options());
                foreach (GeometryObject go in geoel)
                {
                    if (go is GeometryInstance)
                    {
                        GeometryInstance gi = go as GeometryInstance;
                        GeometryElement gl = gi.GetInstanceGeometry();
                        if (gl != null)
                        {
                            foreach (GeometryObject go2 in gl)
                            {
                                if (go2 is PolyLine)
                                {
                                    GraphicsStyle gstyle = Doc.GetElement(go2.GraphicsStyleId) as GraphicsStyle;
                                    string layer = gstyle.GraphicsStyleCategory.Name;
                                    if (!layersname.Contains(layer))
                                    {
                                        layersname.Add(layer);
                                    }

                                    plines.Add(go2 as PolyLine);
                                }
                            }

                        }
                    }
                }
            }
            else
            {
                TaskDialog.Show("error", "can't find cad import");
            }
        }


        public void ExecuteCreateFoundationCommand(object parameter)
        {
            try
            {
                //Assign selected Level 
                IList<Level> levelsname = new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                Level collevel = null;
                foreach (Level level in levelsname)
                {
                    if (level.Name == _selectedlevel)
                    {
                        collevel = level;
                    }
                }


                //Assign Selected Layer
                string selectedlay = Selectedlayer;

                foreach (PolyLine pl in plines)
                {
                    GraphicsStyle gstyle = Doc.GetElement(pl.GraphicsStyleId) as GraphicsStyle;
                    string layer = gstyle.GraphicsStyleCategory.Name;

                    if (layer == _selectedlayer)
                    {
                        #region If User Select the Type
                        //Here We Get The Midline to draw the element
                        Outline o = pl.GetOutline();
                        XYZ fpoint = o.MaximumPoint;
                        XYZ spoint = o.MinimumPoint;
                        XYZ midline = MidPoint(fpoint.X, fpoint.Y, fpoint.Z, spoint.X, spoint.Y, spoint.Z);

                        //Here We Create A family For the Element
                        IList<XYZ> points = pl.GetCoordinates();
                        XYZ p1 = points[0];
                        XYZ p2 = points[1];
                        XYZ p3 = points[2];

                        Vector3D v1 = new Vector3D();
                        v1.X = Math.Abs(p2.X - p1.X);
                        v1.Y = Math.Abs(p2.Y - p1.Y);
                        v1.Z = Math.Abs(p2.Z - p1.Z);

                        Vector3D v2 = new Vector3D();
                        v2.X = Math.Abs(p3.X - p2.X);
                        v2.Y = Math.Abs(p3.Y - p2.Y);
                        v2.Z = Math.Abs(p3.Z - p2.Z);


                        IList<Element> footings = new FilteredElementCollector(Doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsElementType().ToElements();
                        string name = null;
                        bool flag = true;

                        double length;
                        double width;
                        if (v1.X > v2.X)
                        {
                            length = Math.Round(SideLength(p1.X, p1.Y, p1.Z, p2.X, p2.Y, p2.Z));
                            width = Math.Round(SideLength(p2.X, p2.Y, p2.Z, p3.X, p3.Y, p3.Z));

                        }
                        else
                        {
                            width = Math.Round(SideLength(p1.X, p1.Y, p1.Z, p2.X, p2.Y, p2.Z));
                            length = Math.Round(SideLength(p2.X, p2.Y, p2.Z, p3.X, p3.Y, p3.Z));
                        }

                        name = length.ToString() + " x " + width.ToString() + " x " + Depth + "mm";
                        double depth = Convert.ToDouble(Depth); 
                        foreach (var item in footings)
                        {
                            if (item.Name == name)
                            {
                                flag = false;
                            }
                        }
                        if (flag)
                        {
                            FamilySymbol familySymbol = new FilteredElementCollector(Doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsElementType().Where(x => x.Name == "1800 x 1200 x 450mm").FirstOrDefault() as FamilySymbol;
                            Family family = familySymbol.Family;
                            Autodesk.Revit.DB.Document familyDoc = Doc.EditFamily(family);
                            if (null == familyDoc)
                            {
                                return;    // could not open a family for edit
                            }

                            FamilyManager familyManager = familyDoc.FamilyManager;
                            if (null == familyManager)
                            {
                                return;  // could not get a family manager
                            }

                            using (Transaction newFamilyTypeTransaction = new Transaction(familyDoc, "Add Type to Family"))
                            {
                                int changesMade = 0;
                                newFamilyTypeTransaction.Start();
                                // add a new type and edit its parameters                           
                                FamilyType newFamilyType = familyManager.NewType(name);

                                if (newFamilyType != null)
                                {
                                    FamilyParameter familyParam = familyManager.get_Parameter("Length");
                                    if (null != familyParam)
                                    {
                                        familyManager.Set(familyParam, width / 304.8);
                                        changesMade += 1;
                                    }

                                    familyParam = familyManager.get_Parameter("Width");
                                    if (null != familyParam)
                                    {
                                        familyManager.Set(familyParam, length / 304.8);
                                        changesMade += 1;
                                    }

                                    familyParam = familyManager.get_Parameter("Foundation Thickness");
                                    if (null != familyParam)
                                    {
                                        familyManager.Set(familyParam, depth / 304.8);
                                        changesMade += 1;
                                    }
                                }

                                if (3 == changesMade)   
                                {
                                    newFamilyTypeTransaction.Commit();
                                }
                                else 
                                {
                                    newFamilyTypeTransaction.RollBack();
                                }
                                if (newFamilyTypeTransaction.GetStatus() != TransactionStatus.Committed)
                                {
                                    return;
                                }
                            }

                            // now update the Revit project with Family which has a new type
                            LoadOpts loadOptions = new LoadOpts();

                            family = familyDoc.LoadFamily(Doc, loadOptions);
                            if (null != family)
                            {
                                // find the new type and assign it to FamilyInstance
                                ISet<ElementId> familySymbolIds = family.GetFamilySymbolIds();
                                foreach (ElementId id in familySymbolIds)
                                {
                                    FamilySymbol familysymbol = family.Document.GetElement(id) as FamilySymbol;
                                    if ((null != familysymbol) && familysymbol.Name == name)
                                    {
                                        using (Transaction changeSymbol = new Transaction(Doc, "Change Symbol Assignment"))
                                        {
                                            changeSymbol.Start();
                                            familySymbol = familysymbol;
                                            changeSymbol.Commit();
                                        }
                                        break;
                                    }
                                }
                            }
                        }


                        IList<Element> Footingss = new FilteredElementCollector(Doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsElementType().ToElements();
                        FamilySymbol fss = null;
                        foreach (Element ele in Footingss)
                        {
                            if (ele.Name == name)
                            {
                                fss = ele as FamilySymbol;
                            }
                        }
                        using (Transaction trans = new Transaction(Doc, "create Isolated Foundations"))
                        {
                            trans.Start();
                            try
                            {
                                if (!fss.IsActive)
                                {
                                    fss.Activate();
                                }
                                Doc.Create.NewFamilyInstance(midline, fss, collevel, StructuralType.Footing);
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show(ex.Message, ex.ToString());

                            }
                            trans.Commit();
                        }
                        #endregion

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public bool CanExecuteCreateFoundationCommand(object parameter)
        {
            return true;
        }

        private static XYZ MidPoint(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            XYZ midPoint = new XYZ((x1 + x2) / 2, (y1 + y2) / 2, (z1 + z2) / 2);
            return midPoint;
        }
        private static double SideLength(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            double length;
            length = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2) + Math.Pow((z2 - z1), 2)) * 304.8;
            return length;
        }
        #endregion
    }
}
