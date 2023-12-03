using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using LevelsCreationExcel.Commands;
using LevelsCreationExcel.RevitCommands;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;

using System.Data.OleDb;

//This means that we can access the static members of that namespace directly, without having to prefix them with the namespace name.
using static LevelsCreationExcel.Utils.ExcelUtils;
using System.Data;

namespace LevelsCreationExcel.ViewModel
{
    public class ExcelWindowViewModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void onPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Fields&Properties
        private string _filePath;
        public string FilePath
        {
            get => _filePath;
            set
            {
                _filePath = value;
                onPropertyChanged(); // Notify property change
            }
        }

        private List<string> _sheetNames;
        public List<string> SheetNames
        {
            get => _sheetNames;
            set { _sheetNames = value; onPropertyChanged(); }
        }

        private string _selectedSheetName;
        public string SelectedSheetName
        {
            get => _selectedSheetName;
            set
            {
                _selectedSheetName = value;
                onPropertyChanged();
            }
        }

        private string _selectedFilePath;
        public string SelectedFilePath
        {
            get { return _selectedFilePath; }
            set
            {
                if (_selectedFilePath != value)
                {
                    _selectedFilePath = value;
                    onPropertyChanged();
                }
            }
        }

        #endregion

        #region Constructor
        public ExcelWindowViewModel()
        {
            OpenExcelCommand = new MyCommand(ExecuteOpenExcel, CanExecuteOpenExcel);
            CreateLevelsCommand = new MyCommand(ExecuteCreateLevels , CanExecuteCreateLevels);
        }
        #endregion

        #region Commands
        public MyCommand OpenExcelCommand { get; set; }
        public MyCommand CreateLevelsCommand { get; set; }
        #endregion

        #region Methods
        public void ExecuteOpenExcel(object parameter)
        {
            FilePath = OpenExcelAndGetFilePath();
            SheetNames = RetriveSheetsNames(FilePath);
        }
        public bool CanExecuteOpenExcel(object parameter)
        {
            return true;
        }

        public void ExecuteCreateLevels(object parameter)
        {
            Autodesk.Revit.DB.Document doc = OpenExcelWindowCommand.Doc;
            UIDocument uidoc = OpenExcelWindowCommand.uIDocument;

            using (Transaction tr = new Transaction(doc, "Draw Level"))
            {
                try
                {
                    tr.Start();
                    //Excel
                    try
                    {
                        // Read the contents of the Excel file into a DataSet
                        var package = new ExcelPackage(FilePath);
                        using (package)
                        {
                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            var sheet = package.Workbook.Worksheets[SelectedSheetName];

                            string firstcell = " ";
                            double secondcell = 0;

                            for (int i = sheet.Dimension.Start.Row +1; i <= sheet.Dimension.End.Row; i++)
                            {
                                try
                                {
                                    var name = sheet.Cells[i, 1].Value;
                                    if (name == null)  // Skip empty rows
                                        continue;
                                    firstcell = Convert.ToString(name);

                                    var elev = sheet.Cells[i, 2].Value;
                                    if (elev == null) // Skip incomplete rows
                                        continue;
                                    secondcell = Convert.ToDouble(elev);

                                    Level level = Level.Create(doc, UnitUtils.ConvertToInternalUnits(secondcell , DisplayUnitType.DUT_MILLIMETERS));
                                    level.Name =firstcell;

                                    // Create the floor plan view
                                    ViewFamilyType floorPlanType = new FilteredElementCollector(doc)
                                        .OfClass(typeof(ViewFamilyType))
                                        .Cast<ViewFamilyType>()
                                        .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.StructuralPlan);

                                    if (floorPlanType != null)
                                    {
                                        ViewPlan floorPlan = ViewPlan.Create(doc, floorPlanType.Id, level.Id);
                                        floorPlan.Name = level.Name + " Structural Plan";
                                    }

                                }
                                catch (Exception Ex)
                                {
                                    MessageBox.Show(Ex.Message);
                                }
                            }

                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message);
                    }


                    tr.Commit();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.Message);
                }
            }
        }
        public bool CanExecuteCreateLevels(object parameter)
        {
            return true;
        }
        #endregion

    }
}
