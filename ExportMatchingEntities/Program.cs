using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsCrmTools.Translator;

namespace ExportMatchingEntities
{
    class Program
    {
        private static string _filePath = null;

        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Log($"Path not specified");
                return;
            }

            var translatedPath = args[0];
            var allTranslatedFiles = Directory.GetFiles(translatedPath);

            ExportSettings settings = new ExportSettings
            {
                ExportAttributes = true,
                ExportBooleans = true,
            };

            //foreach (var translatedFile in allTranslatedFiles)
            //{
            //    settings = GetTranslationSettings(translatedFile);

            //    var e = new Engine();
            //    e.Export(settings, ser)
            //}

        }

        private static void Log(string msg)
        {
            try
            {
                if (string.IsNullOrEmpty(_filePath))
                {
                    _filePath = "Logs\\ImportTranslations_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log";
                }

                Console.WriteLine(msg);
                File.AppendAllText(_filePath, $"{Environment.NewLine}{DateTime.Now.ToString()} {msg}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during logging {ex.Message}");
            }
        }

    }
}
