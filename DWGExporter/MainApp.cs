// (C) Copyright 2019 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software
// in object code form for any purpose and without fee is hereby
// granted, provided that the above copyright notice appears in
// all copies and that both that copyright notice and the limited
// warranty and restricted rights notice below appear in all
// supporting documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK,
// INC. DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL
// BE UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is
// subject to restrictions set forth in FAR 52.227-19 (Commercial
// Computer Software - Restricted Rights) and DFAR 252.227-7013(c)
// (1)(ii)(Rights in Technical Data and Computer Software), as
// applicable.
//

using System;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace Autodesk.Forge.DWGExporter
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainApp : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public void HandleApplicationInitializedEvent(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            DesignAutomationData data = new DesignAutomationData(app, "InputFile.rvt");
            this.ExportDWG(data);
        }

        private void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered...");
            e.Succeeded = true;
            e.Succeeded = this.ExportDWG(e.DesignAutomationData);
        }

        private bool ExportDWG(DesignAutomationData data)
        {
            if (data == null)
                return false;

            Application app = data.RevitApp;
            if (app == null)
                return false;

            string modelPath = data.FilePath;
            if (string.IsNullOrWhiteSpace(modelPath))
                return false;

            var doc = data.RevitDoc;
            if (doc == null)
                return false;

            using (var collector = new FilteredElementCollector(doc))
            {
                LogTrace("Collecting sheets...");

                var exportPath = Path.GetDirectoryName(modelPath);
                var sheetIds = collector.WhereElementIsNotElementType()
                                        .OfClass(typeof(ViewSheet))
                                        .ToElementIds();

                using (var trans = new Transaction(doc, "Export DWG"))
                {
                    LogTrace("Starting the export task...");

                    try
                    {
                        if (trans.Start() == TransactionStatus.Started)
                        {
                            var dwgSettings = ExportDWGSettings.FindByName(doc, "Forge");
                            if (dwgSettings == null)
                            {
                                dwgSettings = ExportDWGSettings.Create(doc, "Forge");
                                trans.Commit();
                            }

                            var exportOpts = dwgSettings.GetDWGExportOptions();
                            exportOpts.MergedViews = true;

                            LogTrace("Exporting...");

                            doc.Export(exportPath, "DA4R", sheetIds, exportOpts);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (trans.HasStarted() == true)
                            trans.RollBack();

                        LogTrace("Error occured");
                        LogTrace(ex.Message);
                        LogTrace(ex.InnerException.Message);
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args) { System.Console.WriteLine(format, args); }
    }
}
