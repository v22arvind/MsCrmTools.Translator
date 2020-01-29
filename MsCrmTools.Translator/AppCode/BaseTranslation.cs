using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Messages;

namespace MsCrmTools.Translator.AppCode
{
    public class TranslationResultEventArgs : EventArgs
    {
        public bool Success { get; set; }
        public string SheetName { get; set; }
        public string Message { get; set; }
    }
    public class BaseTranslation
    {
        public event EventHandler<TranslationResultEventArgs> Result;
        public static bool AllowBlank = false;
        public virtual void OnResult(TranslationResultEventArgs e)
        {
            EventHandler<TranslationResultEventArgs> handler = Result;
            if (handler != null)
            {
                handler(this, e);
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(e.Message))
                {
                    File.AppendAllText("Logs\\ImportTranslations_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log",
                        string.Format("{0}{1} - {2} - {3}", Environment.NewLine, e.SheetName, e.Success, e.Message));
                }
            }
            catch
            {
            }
        }

        public void ProcessMultiple<T>(IOrganizationService service, List<T> requests, string sheetName, int batch = 5)
        {
            ExecuteMultipleRequest mr = new ExecuteMultipleRequest()
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings() { ContinueOnError = true, ReturnResponses = true }
            };

            int counter = 0;

            foreach (var request in requests)
            {
                mr.Requests.Add((object)request as OrganizationRequest);
                counter++;

                if (mr.Requests.Count > batch)
                {
                    ExecuteTheRequest(service, mr, sheetName, counter);

                    mr.Requests.Clear();
                }
            }

            if (mr.Requests.Count > 0)
            {
                ExecuteTheRequest(service, mr, sheetName, counter);
            }
        }

        private void ExecuteTheRequest(IOrganizationService service, ExecuteMultipleRequest mr, string sheetName, int counter)
        {
            try
            {
                HandleResponse(service.Execute(mr) as ExecuteMultipleResponse, sheetName);

                OnResult(new TranslationResultEventArgs
                {
                    Success = true,
                    SheetName = sheetName,
                    Message = $"Processed {counter} records"
                });
            }
            catch (Exception ex)
            {
                OnResult(new TranslationResultEventArgs
                {
                    Success = false,
                    SheetName = sheetName,
                    Message = $"Error during executing multiple request : {ex.Message}"
                });
            }
        }

        private void HandleResponse(ExecuteMultipleResponse resp, string sheetName)
        {
            if (resp.IsFaulted)
            {
                foreach (var response in resp.Responses)
                {
                    if (response.Fault != null)
                    {
                        OnResult(new TranslationResultEventArgs
                        {
                            Success = false,
                            SheetName = sheetName,
                            Message = $"{response.Fault.Message}"
                        });
                    }
                }
            }
        }
    }
}
