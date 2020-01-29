using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using MsCrmTools.Translator;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using System.IO;
using OfficeOpenXml;

namespace ImportTranslations
{
    class Program
    {
        private static string _filePath = null;

        static void Main(string[] args)
        {

            string directory = string.Empty;
            int? lcidToProcess = null;
            string connStr = null;
            bool export = false;

            Directory.CreateDirectory("Logs");
            Log($"----------------------------------------------");
            Log($"Running the translation import in batch mode");
            Log($"----------------------------------------------");
            try
            {
                //get arguments
                //if (args.Length < 1)
                //{
                //    directory = @"C:\Users\arvind-v\Documents\UNHCR\Data\Translation\XrmToolbox\toprocess";
                //    lcidToProcess = 1036;

                //    Log($"Directory not specified! - running test mode for path: {directory}");
                //}
                //else
                //{
                directory = args[0];

                if (args.Length > 1)
                {
                    connStr = args[1];

                    if (ConfigurationManager.ConnectionStrings.Cast<ConnectionStringSettings>()
                            .FirstOrDefault(s => connStr.Equals(s.Name, StringComparison.OrdinalIgnoreCase)) == null)
                    {
                        Log($"Connection string '{connStr}' is not available in config. Aborting ....");
                        return;
                    }
                }

                if (args.Length > 2)
                {
                    if (string.Equals(args[2], "E", StringComparison.OrdinalIgnoreCase))
                    {
                        export = true;
                    }
                }

                if (args.Length > 2)
                {
                    if (string.Equals(args[2], "E", StringComparison.OrdinalIgnoreCase))
                    {
                        export = true;
                    }
                }


                //if (args.Length > )
                //{
                //    if (int.TryParse(args[2], out int parsedInt))
                //    {
                //        lcidToProcess = parsedInt;
                //    }
                //}
                //}

                Log($"Processing for directory {directory}");

                if (lcidToProcess.HasValue)
                    Log($"Processing for Language {lcidToProcess}");

                if (!Directory.Exists(directory))
                {
                    Log($"Incorrect path! {directory}");
                    return;
                }

                //Compose CRM Service
                CrmServiceClient c =
                    new CrmServiceClient(ConfigurationManager.ConnectionStrings[connStr].ConnectionString);

                Log("Testing CRM service");
                //test service
                var resp = c.Execute(new WhoAmIRequest()) as WhoAmIResponse;
                Log($"Service is working fine. Current user is {resp.UserId}");

                var filesToImport = Directory.GetFiles(directory);
                //move the file to processed folder
                var destDir = Path.Combine(directory, "Processed", DateTime.Now.ToString("yyyyMMddHHmmss"));

                Engine e = new Engine();
                var settings = new ExportSettings();

                foreach (var file in filesToImport)
                {
                    try
                    {
                        if (!export)
                        {
                            Log($"************* Importing File: {file}");

                            e.Import(file, c, new BackgroundWorker(), lcidToProcess);
                        }
                        else
                        {
                            Log($"************* Exporting File: {file}");
                            settings = GetExportSettings(file);
                            e.Export(settings, c, new BackgroundWorker(), false);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogAllError(ex);
                        //Log($"Error processing file {file}. Error Message: {ex.Message}, More Details: {ex.StackTrace}");
                    }

                    if (string.IsNullOrWhiteSpace(file) || !file.EndsWith(".xlsx"))
                        continue;
                    try
                    {

                        Directory.CreateDirectory(destDir);
                        File.Move(file, Path.Combine(destDir, Path.GetFileName(file)));
                    }
                    catch (Exception ex)
                    {
                        Log($"Error in moving file {file}. Error Message: {ex.Message}, More Details: {ex.StackTrace}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Exception occurred! - Message: {ex.Message}");
                Log(ex.StackTrace);
            }
            Log($"****************************");
            Log($"Translation Import Complete");
            Log($"****************************");

        }

        private static void LogAllError(Exception ex)
        {
            Log($"Error processing file. Error Message: {ex.Message}, Source: {ex.Source}");

            if (ex.InnerException != null)
                LogAllError(ex.InnerException);
        }

        private static ExportSettings GetExportSettings(string file)
        {
            var s = new ExportSettings();
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var doc = new ExcelPackage(stream);

                //s.ExportForms = doc.Workbook.Worksheets.Any(x => x.Name.StartsWith("Forms "));
                s.ExportAttributes = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Attributes", StringComparison.OrdinalIgnoreCase));
                s.ExportBooleans = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Booleans", StringComparison.OrdinalIgnoreCase));
                s.ExportCharts = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Charts", StringComparison.OrdinalIgnoreCase));
                s.ExportCustomizedRelationships = doc.Workbook.Worksheets.Any(x => x.Name.StartsWith("Relationships", StringComparison.OrdinalIgnoreCase));
                s.ExportDashboards = doc.Workbook.Worksheets.Any(x => x.Name.StartsWith("Dashboards ", StringComparison.OrdinalIgnoreCase));
                s.ExportDescriptions = true;
                s.ExportEntities = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Entities", StringComparison.OrdinalIgnoreCase));
                s.ExportFormFields = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Forms Fields", StringComparison.OrdinalIgnoreCase));
                s.ExportForms = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Forms", StringComparison.OrdinalIgnoreCase));
                s.ExportFormSections = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Forms Sections", StringComparison.OrdinalIgnoreCase));
                s.ExportFormTabs = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Forms Tabs", StringComparison.OrdinalIgnoreCase));
                s.ExportGlobalOptionSet = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Global OptionSets", StringComparison.OrdinalIgnoreCase));
                s.ExportNames = true;
                s.ExportOptionSet = doc.Workbook.Worksheets.Any(x => x.Name.Equals("OptionSets", StringComparison.OrdinalIgnoreCase));
                s.ExportSiteMap = doc.Workbook.Worksheets.Any(x => x.Name.StartsWith("SiteMap ", StringComparison.OrdinalIgnoreCase));
                s.ExportViews = doc.Workbook.Worksheets.Any(x => x.Name.Equals("Views", StringComparison.OrdinalIgnoreCase));
                s.FilePath = Path.Combine(Path.GetDirectoryName(file), "MatchingExport", Path.GetFileName(file));
                s.Entities = GetAllEntitiesInFile(doc.Workbook.Worksheets.FirstOrDefault(x =>
                    x.Name.Equals("Entities", StringComparison.OrdinalIgnoreCase)));
            }

            Log($"Exporting file {s.FilePath} for entities {string.Join(",", s.Entities)}");

            return s;
        }

        private static IEnumerable<string> GetAllEntitiesInFile(ExcelWorksheet excelWorksheet)
        {
            Log($"Processing sheet: {excelWorksheet.Name}");
            var entities = new List<string>();

            for (int i = 1; i < excelWorksheet.Dimension.Rows; i++)
            {
                var ent = ZeroBasedSheet.Cell(excelWorksheet, i, 1).Value.ToString();

                if (!entities.Contains(ent))
                    entities.Add(ent);
            }

            return entities;
        }

        private static void Log(string msg)
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                _filePath = "Logs\\ImportTranslations_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log";
            }

            Console.WriteLine(msg);
            File.AppendAllText(_filePath, $"{Environment.NewLine}{DateTime.Now.ToString()} {msg}");
        }
    }
}
