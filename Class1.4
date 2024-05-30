using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using Autodesk.Revit.ApplicationServices;

namespace ClassLibrary2
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
                Application app = uiApp.Application;

                // Open the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    // Get all files in the folder
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    foreach (string filePath in files)
                    {
                        Document doc = null;
                        bool saveChanges = false;

                        try
                        {
                            // Open the Revit file
                            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                            OpenOptions openOptions = new OpenOptions();
                            doc = app.OpenDocumentFile(modelPath, openOptions);

                            // If the document was successfully opened
                            if (doc != null)
                            {
                                // Write success message to log file
                                writer.WriteLine(DateTime.Now.ToString() + ": Successfully opened Revit file: " + filePath);

                                // Perform any necessary processing here
                                // Check for the panel-circuit mismatch
                                bool mismatchFound = CheckPanelCircuitMismatch(doc); // Implement this method based on your logic

                                if (mismatchFound)
                                {
                                    // Disconnect the panel automatically and log the action
                                    using (Transaction trans = new Transaction(doc, "Disconnect Panel from Circuit"))
                                    {
                                        trans.Start();
                                        try
                                        {
                                            DisconnectPanelFromCircuit(doc, writer); // Implement this method based on your logic
                                            saveChanges = true;
                                            writer.WriteLine(DateTime.Now.ToString() + ": Disconnected panel from circuit for file: " + filePath);
                                            trans.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                            trans.RollBack();
                                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred during the transaction for file " + filePath + ": " + ex.Message);
                                            writer.WriteLine(ex.StackTrace);
                                        }
                                    }
                                }
                                else
                                {
                                    // If no mismatch was found, mark saveChanges to true
                                    saveChanges = true;
                                }

                                if (saveChanges)
                                {
                                    // Save the document outside of any active transaction
                                    SaveDocument(doc, writer);
                                }

                                // Export to IFC using Autodesk.IFC.Export.UI
                                ExportToIFC(doc, filePath, writer);
                            }
                            else
                            {
                                HandleFileOpeningError(filePath, writer);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Write detailed error message to log file
                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred with file " + filePath + ": " + ex.Message);
                            writer.WriteLine(ex.StackTrace);

                            // Automatically handle the error by disconnecting the panel
                            if (doc != null)
                            {
                                try
                                {
                                    using (Transaction trans = new Transaction(doc, "Disconnect Panel from Circuit on Error"))
                                    {
                                        trans.Start();
                                        DisconnectPanelFromCircuit(doc, writer); // Implement this method based on your logic
                                        saveChanges = true;
                                        writer.WriteLine(DateTime.Now.ToString() + ": Automatically disconnected panel from circuit for file: " + filePath);
                                        trans.Commit();
                                    }
                                }
                                catch (Exception innerEx)
                                {
                                    writer.WriteLine(DateTime.Now.ToString() + ": An additional error occurred while trying to disconnect the panel for file " + filePath + ": " + innerEx.Message);
                                    writer.WriteLine(innerEx.StackTrace);
                                }
                            }
                        }
                        finally
                        {
                            // Ensure the document is closed if it was opened
                            if (doc != null && doc.IsValidObject)
                            {
                                doc.Close(false); // Close without saving changes again
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Log the operation cancellation
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + ": The operation was cancelled by the user.");
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(DateTime.Now.ToString() + ": An error occurred: " + ex.Message);
                }
            }

            return Result.Succeeded;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer)
        {
            // Log the error message to the log file
            writer.WriteLine(DateTime.Now.ToString() + ": Failed to open Revit file: " + filePath);
        }

        private void SaveDocument(Document doc, StreamWriter writer)
        {
            try
            {
                doc.Save();
                writer.WriteLine(DateTime.Now.ToString() + ": Document saved successfully.");
            }
            catch (Exception ex)
            {
                writer.WriteLine(DateTime.Now.ToString() + ": An error occurred while saving the document: " + ex.Message);
                writer.WriteLine(ex.StackTrace);
            }
        }

        // Implement this method to check for panel-circuit mismatch
        private bool CheckPanelCircuitMismatch(Document doc)
        {
            // Logic to check for mismatches in panel circuits
            return false; // Return true if mismatch found, otherwise false
        }

        // Implement this method to disconnect the panel from the circuit
        private void DisconnectPanelFromCircuit(Document doc, StreamWriter writer)
        {
            // Logic to disconnect panel from circuits
        }

        // Implement this method to export the document to IFC
        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                // Ensure the Autodesk.IFC.Export.UI reference is available
                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    // Customize the IFC export options as needed
                    FileVersion = IFCVersion.IFC2x3CV2, // Set the desired IFC version
                    ExportBaseQuantities = true // Export base quantities
                };

                string ifcFolderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Revit2024 Export"; // Specify the destination folder
                Directory.CreateDirectory(ifcFolderPath); // Create the folder if it doesn't exist

                string ifcFileName = Path.GetFileNameWithoutExtension(filePath) + ".ifc"; // Use the same file name as the Revit file, but with the .ifc extension
                string ifcFilePath = Path.Combine(ifcFolderPath, ifcFileName);

                using (Transaction exportTrans = new Transaction(doc, "Export IFC"))
                {
                    exportTrans.Start();
                    doc.Export(ifcFolderPath, ifcFileName, ifcOptions);
                    exportTrans.Commit();
                }

                writer.WriteLine(DateTime.Now.ToString() + ": IFC export completed for: " + ifcFilePath);
            }
            catch (Exception ex)
            {
                writer.WriteLine(DateTime.Now.ToString() + ": An error occurred during IFC export for file " + filePath + ": " + ex.Message);
                writer.WriteLine(ex.StackTrace);
            }
        }
    }
}
