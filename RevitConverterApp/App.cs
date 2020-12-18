#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace RevitConverterApp
{
    class App : IExternalApplication
    {
        public ExternalEvent ConvertEvent { get; set; }
        public static bool IsProcessing = false;
        public static string EmptyRevitFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Revit Converter", "Empty File.rvt");

        public static string ToProcessFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Revit Converter", "Excel To Convert");
        public static string ProcessedFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Revit Converter", "Done Excel");
        public static string RevitOutputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Revit Converter", "Revit Output");

        public Result OnStartup(UIControlledApplication a)
        {
            CreateFiles();
            ConvertEvent = ExternalEvent.Create(new Converter());
            a.Idling += A_Idling;
            return Result.Succeeded;
        }
        private void A_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            if (IsProcessing)
                return;
            CreateFiles();

            var toProcessFiles = Directory.GetFiles(ToProcessFolder).ToList().Where(o => Path.GetExtension(o) == ".csv").ToList();
            if (toProcessFiles.Count == 0)
                return;
            IsProcessing = true;
            Converter.CSVFileToConvert = toProcessFiles[0];
            ConvertEvent.Raise();
        }

        private static void CreateFiles()
        {
            if (!Directory.Exists(ToProcessFolder))
                Directory.CreateDirectory(ToProcessFolder);
            if (!Directory.Exists(ProcessedFolder))
                Directory.CreateDirectory(ProcessedFolder);
            if (!Directory.Exists(RevitOutputFolder))
                Directory.CreateDirectory(RevitOutputFolder);
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
