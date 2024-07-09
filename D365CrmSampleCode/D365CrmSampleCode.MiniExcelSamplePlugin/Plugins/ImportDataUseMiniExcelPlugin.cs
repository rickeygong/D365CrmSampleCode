using Microsoft.Xrm.Sdk;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365CrmSampleCode.MiniExcelSamplePlugin.Plugins
{
    public class ImportDataUseMiniExcelPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            IOrganizationService serviceAdmin = factory.CreateOrganizationService(null);
            string _recordGuid = context.InputParameters["recordGuid"].ToString();
            string _fileBase64 = context.InputParameters["fileBase64"].ToString();
            string _entityName = context.InputParameters["entityName"].ToString();
            try
            {
                string strResults = string.Empty;
                string outMessage = string.Empty;
                string state = string.Empty;  // Success: 0, Failure: 1
                string msg = string.Empty;

                // Verify
                if (string.IsNullOrEmpty(_recordGuid))
                {
                    state = "1";
                    msg = "Parameter exception, \"recordGuid\" is empty.";
                }
                if (string.IsNullOrEmpty(_fileBase64))
                {
                    state = "1";
                    msg = "Parameter exception, \"fileBase64\" is empty.";
                }
                if (string.IsNullOrEmpty(_entityName))
                {
                    state = "1";
                    msg = "Parameter exception, \"entityName\" is empty.";
                }
                if (state == "1")
                {
                    // Output
                    context.OutputParameters["state"] = state;
                    context.OutputParameters["msg"] = msg;
                    return;
                }

                if (_entityName == "gdh_invoice")
                {
                    ImportInvoiceDetail(base64Excel: _fileBase64, recordGuid: _recordGuid, entityName: _entityName, service: serviceAdmin, ref state, ref msg);
                }

                // Output
                context.OutputParameters["state"] = state;
                context.OutputParameters["msg"] = msg;
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
        private void ImportInvoiceDetail(string base64Excel, string recordGuid, string entityName, IOrganizationService service, ref string state, ref string msg)
        {
            byte[] excelBytes = Convert.FromBase64String(base64Excel);
            List<string> verifyStringList = new List<string>();
            using (MemoryStream stream = new MemoryStream(excelBytes))
            {
                BaseVerifyHeaders(stream, ref verifyStringList);
                if (verifyStringList.Count > 0)
                {
                    StringBuilder errorMessage = new StringBuilder();
                    int inxex = 1;
                    foreach (string item in verifyStringList.Distinct())
                    {
                        errorMessage.AppendLine($"{inxex}.{item}");
                        inxex++;
                    }
                    errorMessage.AppendLine("Import exception, please see the description above");
                    state = "1";
                    msg = errorMessage.ToString();
                    return;
                }

                IEnumerable<ImportInvoiceDateilTemplate> invoices = ReadExcelFile(stream).Where(p => p.Date != null && p.Type != null && p.Amount != null && p.Remark != null);
                if (invoices.Count() > 0)
                {
                    foreach (ImportInvoiceDateilTemplate invoice in invoices)
                    {
                        Entity tagEntity = new Entity("gdh_invoice", Guid.Parse(recordGuid));
                        Entity create_InvoiceDetail = new Entity("gdh_invoice_detail");
                        create_InvoiceDetail["gdh_related_invoice"] = tagEntity.ToEntityReference();
                        create_InvoiceDetail["gdh_date"] = invoice.Date;
                        if (invoice.Amount != null)
                        {
                            create_InvoiceDetail["gdh_amount"] = new Money((decimal)invoice.Amount);
                        }
                        if (invoice.Type.Trim() == "T1")
                        {
                            create_InvoiceDetail["gdh_type"] = new OptionSetValue(800000000);
                        }
                        else if (invoice.Type.Trim() == "T2")
                        {
                            create_InvoiceDetail["gdh_type"] = new OptionSetValue(800000001);
                        }
                        else if (invoice.Type.Trim() == "T3")
                        {
                            create_InvoiceDetail["gdh_type"] = new OptionSetValue(800000002);
                        }
                        create_InvoiceDetail["gdh_remark"] = invoice.Remark;
                        service.Create(create_InvoiceDetail);
                    }
                }
                state = "0";
                msg = "";

            }
        }
        private static IEnumerable<ImportInvoiceDateilTemplate> ReadExcelFile(Stream stream)
        {
            IEnumerable<ImportInvoiceDateilTemplate> query = stream.Query<ImportInvoiceDateilTemplate>();
            foreach (ImportInvoiceDateilTemplate row in query)
            {
                yield return row;
            }
        }
        private static void BaseVerifyHeaders(Stream stream, ref List<string> verifyStringList)
        {
            List<dynamic> rows = stream.Query().ToList();
            if (rows.Count > 0)
            {
                if (rows[0].A != "Date" || rows[0].B != "Type" || rows[0].C != "Amount" || rows[0].D != "Remark")
                {
                    verifyStringList.Add("Warning! The imported file is not a standard template (the order or names of the column headers are incorrect), please check.");
                }
            }
            else
            {
                verifyStringList.Add("Warning! Column headers are empty.");
            }
        }
    }
    public class ImportInvoiceDateilTemplate
    {
        [ExcelColumnName("Type")]
        public string Type { get; set; }
        [ExcelColumnName("Amount")]
        public decimal? Amount { get; set; }
        [ExcelColumnName("Remark")]
        public string Remark { get; set; }
        [ExcelColumnName("Date")]
        public DateTime Date { get; set; }
    }
}
