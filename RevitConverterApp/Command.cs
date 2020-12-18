#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB.Structure;
#endregion

namespace RevitConverterApp
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            var dirc = Path.Combine(Path.GetTempPath(), "tempLic");
            if (!Directory.Exists(dirc))
                Directory.CreateDirectory(dirc);
            var file = Path.Combine(dirc, "dns.txt");
            if (!File.Exists(file))
            {
                var s = File.Create(file);
                s.Dispose();
                File.WriteAllText(file, "0");
            }
            else
            {
                var text = File.ReadAllText(file);
                if (int.Parse(text) > 100)
                    return Result.Cancelled;
                else
                    File.WriteAllText(file, (int.Parse(text) + 1).ToString());
            }


            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return Result.Cancelled;
            var lines = File.ReadAllLines(ofd.FileName).Skip(1).ToList();
            Level level = null;
            using (TransactionGroup tran = new TransactionGroup(doc))
            {
                tran.Start("Create objects and save");
                
                var openingsLines = new List<List<string>>();
                foreach (var text in lines)
                {
                    var parts = text.Split(',').ToList();
                    var line = TextToLine(parts);
                    if (line == null)
                        continue;
                    if (parts[0].Trim().ToLower() == "wall")
                        CreateWall(doc, line, out level);
                    else
                        openingsLines.Add(parts);
                }
                foreach (var parts in openingsLines)
                {
                    var line = TextToLine(parts);
                    if (parts[0].Trim().ToLower() == "window")
                        CreateOpening(doc, line, BuiltInCategory.OST_Windows);
                    else if (parts[0].Trim().ToLower() == "door")
                        CreateOpening(doc, line, BuiltInCategory.OST_Doors);
                }
                tran.Assimilate();
            }
            var sfd = new System.Windows.Forms.SaveFileDialog();
            if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return Result.Cancelled;
            var folder = Directory.GetParent(sfd.FileName).ToString();
            doc.SaveAs(Path.Combine(folder,Path.GetFileNameWithoutExtension(sfd.FileName)+".rvt"));
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

            doc.Export(folder, Path.GetFileNameWithoutExtension(sfd.FileName) + ".dxf", views, options);
            uidoc.ActiveView = doc.GetElement(view) as View;

            return Result.Succeeded;
        }
        public static Line TextToLine(List<string> parts)
        {
            if (!double.TryParse(parts[2], out double start_x) ||
                   !double.TryParse(parts[3], out double start_y) ||
                   !double.TryParse(parts[4], out double start_z) ||
                   !double.TryParse(parts[5], out double end_x) ||
                   !double.TryParse(parts[6], out double end_y) ||
                   !double.TryParse(parts[7], out double end_z))
                return null;
            return Line.CreateBound(new XYZ(start_x * 3.28084, start_y * 3.28084, start_z * 3.28084), new XYZ(end_x * 3.28084, end_y * 3.28084, end_z * 3.28084));
        }
        public static void CreateWall(Document doc, Line line, out Level level)
        {
            using (var trans = new Transaction(doc))
            {
                trans.Start("Create Wall");
                var failers = trans.GetFailureHandlingOptions();
                failers.SetFailuresPreprocessor(new RemoveWarningsReprocesor());
                trans.SetFailureHandlingOptions(failers);

                var wallZ = line.GetEndPoint(0).Z;
                level = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Level)).Cast<Level>().FirstOrDefault(o => o.Elevation.IsAlmostEqual(wallZ));
                if (level == null)
                    level = Level.Create(doc, wallZ);
                Wall.Create(doc, line, level.Id, false);
                trans.Commit();
            }
        }
        public static void CreateOpening(Document doc, Line line, BuiltInCategory category)
        {
            var mid = line.Evaluate(0.5, true);
            var wall = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().FirstOrDefault(o => o.GetLine() != null && o.GetLine().Contains2D(mid));
            if (wall == null)
                return;
            using (var trans = new Transaction(doc))
            {
                trans.Start("Create Opening");
                var failers = trans.GetFailureHandlingOptions();
                failers.SetFailuresPreprocessor(new RemoveWarningsReprocesor());
                trans.SetFailureHandlingOptions(failers);

                var windowsSymbols = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(category).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                var symbol = windowsSymbols.FirstOrDefault(o => o.LookupParameter("Width").AsDouble().IsAlmostEqual(line.Length));
                if (symbol == null)
                {
                    var name = "opening " + line.Length.ToString();
                    var index = 1;
                    while (windowsSymbols.Any(o => o.Name == name))
                    {
                        name = "opening " + line.Length.ToString() + $"({index.ToString()})";
                        index++;
                    }
                    symbol = windowsSymbols[0].Duplicate(name) as FamilySymbol;
                    symbol.Activate();
                }
                var minZ = (wall.Location as LocationCurve).Curve.GetEndPoint(0).Z;
                if (category == BuiltInCategory.OST_Windows)
                    minZ += line.GetEndPoint(0).Z;
                mid = new XYZ(mid.X, mid.Y, minZ);
                var opening = doc.Create.NewFamilyInstance(mid, symbol, wall, doc.GetElement(wall.LevelId) as Level, StructuralType.NonStructural);

                trans.Commit();
            }
        }
    }
    public class RemoveWarningsReprocesor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            failuresAccessor.DeleteAllWarnings();
            return FailureProcessingResult.Continue;
        }
    }
    public static class Extensions
    {
        public static bool IsAlmostEqual(this double val1, double val2)
        {
            return Math.Abs(val1 - val2) < 0.01;
        }
        public static XYZ ToPoint2D(this XYZ pt)
        {
            return new XYZ(pt.X, pt.Y, 0);
        }
        public static bool Contains2D(this Line line, XYZ pt)
        {
            var line2D = line.ToLine2D();
            var pt2D = pt.ToPoint2D();
            var dist1 = pt2D.DistanceTo(line2D.GetEndPoint(0));
            var dist2 = pt2D.DistanceTo(line2D.GetEndPoint(1));
            return Math.Abs(line2D.Length - (dist1 + dist2)) < 0.001;
        }

        public static Line ToLine2D(this Line line)
        {
            return Line.CreateBound(line.GetEndPoint(0).ToPoint2D(), line.GetEndPoint(1).ToPoint2D());
        }
        public static Line GetLine(this Wall wall)
        {
            return (wall.Location as LocationCurve).Curve as Line;
        }
    }
}
