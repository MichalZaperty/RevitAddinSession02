#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel;


#endregion

namespace RevitAddin2
{
    [Transaction(TransactionMode.Manual)]
    public class Session02_Challenge : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Specify excel file
            string excelFile = @"C:\tmp\Session02_Challenge - Copy2.xlsx";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            try
            {

                // Create Levels
                Excel.Worksheet excelWs1 = excelApp.Worksheets[1];
                Excel.Range excelRng1 = excelWs1.UsedRange;
                int rowcount1 = excelRng1.Rows.Count;

                using (Transaction t = new Transaction(doc))

                {
                    t.Start("Bip Bip Bop...");

                    for (int i = 2; i <= rowcount1; i++)
                    {
                        Excel.Range cell1 = excelWs1.Cells[i, 1];
                        Excel.Range cell3 = excelWs1.Cells[i, 3];

                        String levelName = cell1.Value.ToString();
                        double levelElveation = UnitUtils.ConvertToInternalUnits(cell3.Value, DisplayUnitType.DUT_METERS);
                        
                         Level newLevel = Level.Create(doc, levelElveation);
                         newLevel.Name = levelName;  

                    }
                   
                    t.Commit();
                }

                // Create Sheets
                Excel.Worksheet excelWs2 = excelApp.Worksheets[2];
                Excel.Range excelRng2 = excelWs2.UsedRange;
                int rowcount2 = excelRng2.Rows.Count;

                using (Transaction t = new Transaction(doc))

                {
                    t.Start("Bip Bip Bop...");

                    for (int i = 2; i <= rowcount2; i++)
                    {
                        Excel.Range cell1 = excelWs2.Cells[i, 1];
                        Excel.Range cell2 = excelWs2.Cells[i, 2];

                        String sheetName = cell2.Value.ToString();
                        String sheetNumber = cell1.Value.ToString();

                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                        collector.WhereElementIsElementType();

                        ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                        curSheet.Name = sheetName;
                        curSheet.SheetNumber = sheetNumber;
                        
                    }

                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("HALL 9000:", ex.ToString());

            }
            finally
            { 
            excelWb.Close();
            excelApp.Quit();
            
            }
       
            return Result.Succeeded;
        }
    }
}
