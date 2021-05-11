using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Aliquis.ReportSender
{
    public class ReportSender : IPlugin
    {
        #region Secure/Unsecure Configuration Setup
        private string _secureConfig = null;
        private string _unsecureConfig = null;

        public ReportSender(string unsecureConfig, string secureConfig)
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecureConfig;
        }
        #endregion

        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            try
            {
                var reportid = (string)context.InputParameters["reportid"];
                if (string.IsNullOrEmpty(reportid))
                    throw new Exception("Не задан параметр: reportid");
                var parameters = (string)context.InputParameters["parameters"];
                var format = (string)context.InputParameters["format"];
                if (string.IsNullOrEmpty(format))
                    throw new Exception("Не задан параметр: format");
                var userName = (string)context.InputParameters["userName"];
                if (string.IsNullOrEmpty(userName))
                    throw new Exception("Не задан параметр: userName");
                var password = (string)context.InputParameters["password"];
                if (string.IsNullOrEmpty(password))
                    throw new Exception("Не задан параметр: password");
                var domain = (string)context.InputParameters["domain"];
                if (string.IsNullOrEmpty(domain))
                    throw new Exception("Не задан параметр: domain");
                var url = (string)context.InputParameters["url"];
                if (string.IsNullOrEmpty(url))
                    throw new Exception("Не задан параметр: url");
                var toUserId = (string)context.InputParameters["toUserId"];
                if (string.IsNullOrEmpty(toUserId))
                    throw new Exception("Не задан параметр: toUserId");
                var subject = (string)context.InputParameters["subject"];
                if (string.IsNullOrEmpty(subject))
                    throw new Exception("Не задан параметр: subject");
                Entity target = new Entity("report", new Guid(reportid));
                var renderer = new Renderer(url, userName, password, domain);
                var reportColumns = new ColumnSet("reportnameonsrs", "languagecode", "defaultfilter");
                var report = service.Retrieve(target.LogicalName, target.Id, reportColumns);
                var rendered = renderer.RenderReport(report, format, parameters);
                var ext = format == "EXCEL" ? "xlsx" : format == "WORD" ? "docx" : format;
                EmailSender.SendEmail(service, context.UserId, new Guid(toUserId), subject, subject, Convert.ToBase64String(rendered), ext);
                context.OutputParameters["result"] = "OK!";
            }
            catch (Exception ex)
            {
                context.OutputParameters["hasError"] = true;
                context.OutputParameters["result"] = $"{ex.Message} {ex.StackTrace}";
            }
        }

    }
}
