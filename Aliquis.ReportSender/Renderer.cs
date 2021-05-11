using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace Aliquis.ReportSender
{
    public class Renderer
    {
        private readonly CredentialCache _credentialCache;
        private readonly string _url;
        public Renderer(string url, string userName, string password, string doamin)
        {
            _url = url;
            _credentialCache = new CredentialCache
            {
                { new Uri(url), "NTLM", new NetworkCredential(userName, password, doamin) }
            };
        }

        public byte[] RenderReport(Entity report, string format, string parameters)
        {
            var session = GetReportSession(report, parameters);
            format = format.ToUpper();
            if (format == "WORD" || format == "EXCEL")
                format = format + "OPENXML";
            return GetResponse(GetRequest("GET", $"/Reserved.ReportViewerWebControl.axd?ReportSession={session.Item1}&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID={session.Item2}&OpType=Export&FileName=&ContentDisposition=OnlyHtmlInline&Format={format}"));
        }

        private Tuple<string, string> GetReportSession(Entity report, string parameters)
        {
            var name = report.GetAttributeValue<string>("reportnameonsrs");
            var url = "CRMReports/rsviewer/QuirksReportViewer.aspx";
            var parm = $"id=%7B{report.Id}%7D&uniquename={"ymtest3"}&iscustomreport=true&reportnameonsrs=&reportName={name}&isScheduledReport=false";
            if (!string.IsNullOrEmpty(parameters))
            {
                var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(parameters);
                parm += $"&{string.Join("&", values.Select(kvp => string.Format("p:{0}={1}", kvp.Key, kvp.Value)))}";
            }
            var response = Encoding.UTF8.GetString(GetResponse(GetRequest("POST", url, parm)));
            if (response.Contains("ReportSession=") && response.Contains("ControlID="))
            {
                var sessionId = response.Substring(response.LastIndexOf("ReportSession=") + 14, 24);
                var controlId = response.Substring(response.LastIndexOf("ControlID=") + 10, 32);
                return new Tuple<string, string>(sessionId, controlId);
            }
            else
                throw new Exception(parm + "Ошибка при получении сеанса отчета. Скорее всего, это проблема с недопустимыми параметрами отчета." + response);
        }

        private HttpWebRequest GetRequest(string method, string url, string data = null)
        {
            var request = WebRequest.CreateHttp($"{_url}{url}");
            request.Method = method;
            request.Credentials = _credentialCache;
            request.AutomaticDecompression = DecompressionMethods.GZip;
            if (!string.IsNullOrEmpty(data))
            {
                var body = Encoding.ASCII.GetBytes(data);
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = body.Length;
                using (var stream = request.GetRequestStream())
                {
                    stream.Write(body, 0, body.Length);
                }
            }
            return request;
        }

        private byte[] GetResponse(HttpWebRequest request)
        {
            using (var response = request.GetResponse() as HttpWebResponse)
            {
                if (response.ResponseUri.PathAndQuery.Contains("errorhandler.aspx"))
                {
                    var queryString = JsonConvert.DeserializeObject<Dictionary<string, string>>(response.ResponseUri.Query);
                    throw new Exception($"Ошибка при выполнении веб-запроса: {request.RequestUri}");
                }
                using (var stream = response.GetResponseStream())
                using (var stream2 = new MemoryStream())
                {
                    stream.CopyTo(stream2);
                    return stream2.ToArray();
                }
            }
        }
    }
}