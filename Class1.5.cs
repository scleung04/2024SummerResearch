using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;

namespace ClassLibrary2
{
    [Transaction(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string folderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Revit Files";
            string logFilePath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\Dynamo_Revit_Log.txt";

            try
            {
                UIApplication uiApp = commandData.Application;
                Application app = uiApp.Application;

                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    string[] files = Directory.GetFiles(folderPath, "*.rvt");

                    foreach (string filePath in files)
                    {
                        Document doc = null;

                        try
                        {
                            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
                            OpenOptions openOptions = new OpenOptions();
                            doc = app.OpenDocumentFile(modelPath, openOptions);

                            if (doc != null)
                            {
                                writer.WriteLine($"{DateTime.Now}: Successfully opened Revit file: {filePath}");

                                CheckPanelCircuitMismatch(doc, writer);

                                DisconnectPanelFromCircuit(doc, writer);

                                SaveDocument(doc, writer);

                                ExportToIFC(doc, filePath, writer);
                            }
                            else
                            {
                                HandleFileOpeningError(filePath, writer);
                            }
                        }
                        catch (Exception ex)
                        {
                            writer.WriteLine($"{DateTime.Now}: An error occurred with file {filePath}: {ex.Message}");
                            writer.WriteLine(ex.StackTrace);
                        }
                        finally
                        {
                            if (doc != null && doc.IsValidObject)
                            {
                                doc.Close(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: An error occurred: {ex.Message}");
                }
            }

            return Result.Succeeded;
        }

        private void HandleFileOpeningError(string filePath, StreamWriter writer)
        {
            writer.WriteLine($"{DateTime.Now}: Failed to open Revit file: {filePath}");
        }

        private void SaveDocument(Document doc, StreamWriter writer)
        {
            try
            {
                doc.Save();
                writer.WriteLine($"{DateTime.Now}: Document saved successfully.");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{DateTime.Now}: An error occurred while saving the document: {ex.Message}");
                writer.WriteLine(ex.StackTrace);
            }
        }

        private void CheckPanelCircuitMismatch(Document doc, StreamWriter writer)
        {
            var panels = new FilteredElementCollector(doc)
                         .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                         .OfClass(typeof(FamilyInstance))
                         .WhereElementIsNotElementType()
                         .OfType<FamilyInstance>()
                         .ToList();

            foreach (var panel in panels)
            {
                double totalLoad = panel.LookupParameter("Total Connected Load")?.AsDouble() ?? 0;
                double panelCapacity = panel.LookupParameter("Panel Capacity")?.AsDouble() ?? 0;

                if (totalLoad > panelCapacity)
                {
                    writer.WriteLine($"{DateTime.Now}: Mismatch found for panel {panel.Name}. Total load exceeds capacity.");
                }
            }
        }

        private void DisconnectPanelFromCircuit(Document doc, StreamWriter writer)
        {
            var panels = new FilteredElementCollector(doc)
                         .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                         .OfClass(typeof(FamilyInstance))
                         .WhereElementIsNotElementType()
                         .OfType<FamilyInstance>()
                         .ToList();

            foreach (var panel in panels)
            {
                var electricalSystems = panel.MEPModel?.GetElectricalSystems();
                if (electricalSystems != null)
                {
                    foreach (var electricalSystem in electricalSystems)
                    {
                        using (var trans = new Transaction(doc, "Disconnect Panel from Circuit"))
                        {
                            trans.Start();
                            foreach (var connector in electricalSystem.ConnectorManager.Connectors.OfType<Connector>())
                            {
                                foreach (var connectedConnector in connector.AllRefs.OfType<Connector>())
                                {
                                    connector.DisconnectFrom(connectedConnector);
                                    writer.WriteLine($"{DateTime.Now}: Disconnected connector {connector.Owner.Id} from {connectedConnector.Owner.Id}");
                                }
                            }
                            if (ShouldDeleteElement(electricalSystem))
                            {
                                doc.Delete(electricalSystem.Id);
                                writer.WriteLine($"{DateTime.Now}: Deleted electrical system: {electricalSystem.Id.IntegerValue}");
                            }
                            trans.Commit();
                        }
                    }
                }
            }
        }

        private bool ShouldDeleteElement(Element element)
        {
            if (element is ElectricalSystem electricalSystem)
            {
                List<Connector> loadConnectors = new List<Connector>();

                foreach (Element e in electricalSystem.Elements)
                {
                    if (e is FamilyInstance familyInstance && familyInstance.MEPModel != null)
                    {
                        var connectors = familyInstance.MEPModel.ConnectorManager.Connectors
                            .OfType<Connector>()
                            .Where(c => c.IsConnected);

                        loadConnectors.AddRange(connectors);
                    }
                }

                return !loadConnectors.Any();
            }

            return false;
        }

        private void ExportToIFC(Document doc, string filePath, StreamWriter writer)
        {
            try
            {
                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    FileVersion = IFCVersion.IFC2x3CV2,
                    ExportBaseQuantities = true
                };
                string ifcFolderPath = "C:\\Users\\scleu\\Downloads\\2024 Summer Research\\R2024";
                Directory.CreateDirectory(ifcFolderPath);

                string ifcFileName = Path.GetFileNameWithoutExtension(filePath) + ".ifc";
                string ifcFilePath = Path.Combine(ifcFolderPath, ifcFileName);

                using (Transaction exportTrans = new Transaction(doc, "Export IFC"))
                {
                    exportTrans.Start();
                    doc.Export(ifcFolderPath, ifcFileName, ifcOptions);
                    exportTrans.Commit();
                }

                writer.WriteLine($"{DateTime.Now}: IFC export completed for: {ifcFilePath}");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{DateTime.Now}: An error occurred during IFC export for file {filePath}: {ex.Message}");
                writer.WriteLine(ex.StackTrace);
            }
        }
    }
}
