using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;
using System.IO;

namespace ClassLibrary1
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Path to the folder containing Revit files
            string folderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Revit Files";
            // Path to the log file
            string logFilePath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Dynamo_Revit_Log.txt";

            try
            {
                // Get the current Revit application
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Application app = uiApp.Application;

                // Open the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    // Get all files in the folder
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    foreach (string filePath in files)
                    {
                        try
                        {
                            // Open the Revit file
                            Document doc = app.OpenDocumentFile(filePath);

                            // If the document was successfully opened
                            if (doc != null)
                            {
                                TaskDialog.Show("Success", "Successfully opened Revit file: " + filePath);
                                // Write success message to log file
                                writer.WriteLine(DateTime.Now.ToString() + ": Successfully opened Revit file: " + filePath);
                            }
                            else
                            {
                                TaskDialog.Show("Error", "Failed to open Revit file: " + filePath);
                                // Write error message to log file
                                writer.WriteLine(DateTime.Now.ToString() + ": Failed to open Revit file: " + filePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", "An error occurred: " + ex.Message);
                            // Write error message to log file
                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred: " + ex.Message);
                        }
                    }
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