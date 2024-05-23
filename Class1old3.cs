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
        private string userChoice = null;

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
                                TaskDialog.Show("Success", "Successfully opened Revit file: " + filePath);
                                // Write success message to log file
                                writer.WriteLine(DateTime.Now.ToString() + ": Successfully opened Revit file: " + filePath);

                                // Perform any necessary processing here
                                // Check for the panel-circuit mismatch
                                bool mismatchFound = CheckPanelCircuitMismatch(doc); // Implement this method based on your logic

                                if (mismatchFound)
                                {
                                    // Show custom TaskDialog with options
                                    TaskDialog dialog = new TaskDialog("Panel-Circuit Mismatch");
                                    dialog.MainInstruction = "The panel no longer matches the properties for the Circuit.";
                                    dialog.MainContent = "What would you like to do?";
                                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Disconnect the panel from the circuit");
                                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Update properties to match");
                                    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Skip this issue");
                                    dialog.CommonButtons = TaskDialogCommonButtons.Close;
                                    dialog.DefaultButton = TaskDialogResult.Close;

                                    TaskDialogResult result = dialog.Show();

                                    // Perform actions based on user's choice
                                    using (Transaction trans = new Transaction(doc, "Update Revit Document"))
                                    {
                                        trans.Start();
                                        try
                                        {
                                            switch (result)
                                            {
                                                case TaskDialogResult.CommandLink1:
                                                    DisconnectPanelFromCircuit(doc); // Implement this method based on your logic
                                                    saveChanges = true;
                                                    break;
                                                case TaskDialogResult.CommandLink2:
                                                    UpdatePanelCircuitProperties(doc); // Implement this method based on your logic
                                                    saveChanges = true;
                                                    break;
                                                case TaskDialogResult.CommandLink3:
                                                    // Skip: Do nothing
                                                    break;
                                            }
                                            trans.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                            trans.RollBack();
                                            TaskDialog.Show("Error", "An error occurred during the transaction: " + ex.Message);
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
                                HandleFileOpeningError(filePath, writer, ref userChoice);
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Error", "An error occurred: " + ex.Message);
                            // Write detailed error message to log file
                            writer.WriteLine(DateTime.Now.ToString() + ": An error occurred with file " + filePath + ": " + ex.Message);
                            writer.WriteLine(ex.StackTrace);

                            // Handle the error based on user choice
                            HandleFileOpeningError(filePath, writer, ref userChoice);
                        }
                        finally
                        {
                            // Ensure the document is closed if it was opened
                            doc?.Close(false); // Close without saving changes again
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                TaskDialog.Show("Operation Cancelled", "The operation was cancelled by the user.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred: " + ex.Message);
            }

            return Result.Succeeded;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer, ref string userChoice)
        {
            if (userChoice == null)
            {
                // Show custom TaskDialog with options
                TaskDialog dialog = new TaskDialog("Error");
                dialog.MainInstruction = "Failed to open Revit file: " + filePath;
                dialog.MainContent = "What would you like to do?";
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Retry");
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Skip");
                dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
                dialog.CommonButtons = TaskDialogCommonButtons.Close;
                dialog.DefaultButton = TaskDialogResult.Close;

                TaskDialogResult result = dialog.Show();

                // Perform actions based on user's choice
                switch (result)
                {
                    case TaskDialogResult.CommandLink1:
                        userChoice = "Retry";
                        break;
                    case TaskDialogResult.CommandLink2:
                        userChoice = "Skip";
                        break;
                    case TaskDialogResult.CommandLink3:
                        userChoice = "Cancel";
                        throw new OperationCanceledException("Operation cancelled by the user.");
                }
            }

            // Handle the user's choice
            switch (userChoice)
            {
                case "Retry":
                    // Retry: Do nothing, continue to next file
                    break;
                case "Skip":
                    // Skip: Skip this file and continue to next file
                    return; // Skip this iteration of the loop
            }

            // Write error message to log file
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
                TaskDialog.Show("Error", "An error occurred while saving the document: " + ex.Message);
                writer.WriteLine(DateTime.Now.ToString() + ": An error occurred while saving the document: " + ex.Message);
                writer.WriteLine(ex.StackTrace);
            }
        }

        // Implement this method to check for panel-circuit mismatch
        private bool CheckPanelCircuitMismatch(Document doc)
        {
            // Add your logic here to check for mismatch
            return false; // Return true if mismatch found, otherwise false
        }

        // Implement this method to disconnect the panel from the circuit
        private void DisconnectPanelFromCircuit(Document doc)
        {
            // Add your logic here to disconnect the panel from the circuit
        }

        // Implement this method to update panel and circuit properties to match
        private void UpdatePanelCircuitProperties(Document doc)
        {
            // Add your logic here to update properties
        }

        // Implement this method to export the document to IFC
        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                // Ensure the Autodesk.IFC.Export.UI reference is available
                IFCExportOptions ifcOptions = new IFCExportOptions();
                string ifcFilePath = Path.ChangeExtension(filePath, "ifc");

                using (Transaction exportTrans = new Transaction(doc, "Export IFC"))
                {
                    exportTrans.Start();
                    doc.Export(Path.GetDirectoryName(ifcFilePath), Path.GetFileName(ifcFilePath), ifcOptions);
                    exportTrans.Commit();
                }

                TaskDialog.Show("IFC Export", "IFC export completed for: " + ifcFilePath);
                writer.WriteLine(DateTime.Now.ToString() + ": IFC export completed for: " + ifcFilePath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred during IFC export: " + ex.Message);
                writer.WriteLine(DateTime.Now.ToString() + ": An error occurred during IFC export for file " + filePath + ": " + ex.Message);
                writer.WriteLine(ex.StackTrace);
            }
        }
    }
}
