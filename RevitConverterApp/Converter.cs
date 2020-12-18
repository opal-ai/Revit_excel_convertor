using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitConverterApp
{
    public class Converter : IExternalEventHandler
    {
        public static string CSVFileToConvert;
        public void Execute(UIApplication app)
        {

            try
            {
                var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Revit Converter", "Revit Output");
                var file = Path.GetFileNameWithoutExtension(CSVFileToConvert);
                var newRevitFile = Path.Combine(folder, file + ".rvt");
                var index = 1;
                while (File.Exists(newRevitFile))
                {
                    newRevitFile = Path.Combine(folder, file + $"({index.ToString()}).rvt");
                    index++;
                }
                File.Copy(App.EmptyRevitFile, newRevitFile, true);
                var uiDoc = app.OpenAndActivateDocument(newRevitFile);
                var doc = uiDoc.Document;

                var dist = Path.Combine(App.ProcessedFolder, Path.GetFileName(CSVFileToConvert));
                index = 1;
                while (File.Exists(dist))
                {
                    dist = Path.Combine(App.ProcessedFolder, Path.GetFileNameWithoutExtension(CSVFileToConvert) + $"({index.ToString()}).csv");
                    index++;
                }
                File.Move(CSVFileToConvert, dist);
                try
                {
                    File.Delete(CSVFileToConvert);
                }
                catch (Exception)
                {
                }

                var lines = File.ReadAllLines(dist).Skip(1).ToList();
                Level level = null;
                using (TransactionGroup tran = new TransactionGroup(doc))
                {
                    tran.Start("Create objects and save");

                    var openingsLines = new List<List<string>>();
                    foreach (var text in lines)
                    {
                        var parts = text.Split(',').ToList();
                        var line = Command.TextToLine(parts);
                        if (line == null)
                            continue;
                        if (parts[0].Trim().ToLower() == "wall")
                            Command.CreateWall(doc, line, out level);
                        else
                            openingsLines.Add(parts);
                    }
                    foreach (var parts in openingsLines)
                    {
                        var line = Command.TextToLine(parts);
                        if (parts[0].Trim().ToLower() == "window")
                            Command.CreateOpening(doc, line, BuiltInCategory.OST_Windows);
                        else if (parts[0].Trim().ToLower() == "door")
                            Command.CreateOpening(doc, line, BuiltInCategory.OST_Doors);
                    }
                    tran.Assimilate();
                }
                var view = level.FindAssociatedPlanViewId();
                if (view.IntegerValue == -1)
                {
                    using (var trans = new Transaction(doc))
                    {
                        trans.Start("Create View");
                        var viewFamType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().ElementAt(0);
                        view = ViewPlan.Create(doc, viewFamType.Id, level.Id).Id;
                        trans.Commit();
                    }
                }
                var views = new List<ElementId>() { view };
                var options = new DXFExportOptions();
                options.Colors = ExportColorMode.IndexColors;
                options.ExportOfSolids = SolidGeometry.Polymesh;
                options.FileVersion = ACADVersion.R2013;
                options.HideScopeBox = true;
                options.HideUnreferenceViewTags = true;
                options.HideReferencePlane = true;
                options.LayerMapping = "AIA";
                options.LineScaling = LineScaling.PaperSpace;
                options.PropOverrides = PropOverrideMode.ByEntity;
                options.TextTreatment = TextTreatment.Exact;

                doc.Export(folder, file + ".dxf", views, options);

                doc.Save();
                ThreadPool.QueueUserWorkItem(new WaitCallback(CloseDocProc));

                App.IsProcessing = false;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.ToString());
                App.IsProcessing = false;
            }
        }
        static void CloseDocProc(object stateInfo)
        {
            try
            {
                SendKeys.SendWait("^{F4}");
            }
            catch (Exception ex)
            {
            }
        }

        public string GetName()
        {
            return "";
        }
    }
}
