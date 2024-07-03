using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using D365CrmSampleCode.Plugins.MiniExcelSample;
using System.IO;

namespace D365CrmSampleCode.Plugins.Test
{
    [TestClass]
    public class ImportDetailUseMiniExcelTest
    {
        [TestMethod]
        public void ImportDetailUseMiniExcelTestMethod()
        {
            // 读取excel文件，并转为base64
            string filePath = @"D:\MyCode\Mini_excel_sample\MiniExcelSample\ImportInvoiceDetailTemplate.xlsx";
            byte[] fileBytes = File.ReadAllBytes(filePath);
            string base64Excel = Convert.ToBase64String(fileBytes);

            XrmRealContext fakedContext = new XrmRealContext("dev");
            IOrganizationService orgService = fakedContext.GetOrganizationService();
            XrmFakedPluginExecutionContext pluginContext = fakedContext.GetDefaultPluginContext();
            pluginContext.InputParameters = new ParameterCollection
            {
              new KeyValuePair<string, object>("recordGuid","98F83878-1E35-EF11-8409-0017FA0671FA"),
              new KeyValuePair<string, object>("entityName","gdh_invoice"),
              new KeyValuePair<string, object>("fileBase64",base64Excel),
            };
            fakedContext.ExecutePluginWith<ImportDetailUseMiniExcel>(pluginContext);
        }
    }
}
