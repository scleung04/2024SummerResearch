using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;

namespace ClassLibrary1
{
    [Transaction(TransactionMode.Manual)]
    public class OpenRevitFileCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Path to the Revit file you want to open
            string filePath = "C:\\Users\\scleu\\Downloads\\Clinic_A.rvt";

            try
            {
                // Get the current Revit application
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Application app = uiApp.Application;

                // Open the Revit file
                Document doc = app.OpenDocumentFile(filePath);

                // If the document was successfully opened
                if (doc != null)
                {
                    TaskDialog.Show("Success", "Successfully opened Revit file: " + filePath);
                }
                else
                {
                    TaskDialog.Show("Error", "Failed to open Revit file: " + filePath);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred: " + ex.Message);
            }

            return Result.Succeeded;
        }
    }
}

// C:\\Users\\scleu\\Downloads\\Clinic_A.rvt