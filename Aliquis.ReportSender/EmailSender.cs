using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Text;
using System.Web;

namespace Aliquis.ReportSender
{
    public static class EmailSender
    {
        public static void SendEmail(IOrganizationService service, Guid fromUserId, Guid toUserId, string subject, string description, string attachmentBody, string ext)
        {
            Entity fromActivityParty = new Entity("activityparty");
            Entity toActivityParty = new Entity("activityparty");
            fromActivityParty["partyid"] = new EntityReference("systemuser", fromUserId);
            toActivityParty["partyid"] = new EntityReference("systemuser", toUserId);
            Entity email = new Entity("email");
            email["from"] = new Entity[] { fromActivityParty };
            email["to"] = new Entity[] { toActivityParty };
            email["subject"] = subject;
            email["description"] = description;
            email["directioncode"] = true;
            Guid emailId = service.Create(email);
            AddAttachments(service, emailId, subject, attachmentBody, ext);
            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };
            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
        }

        private static void AddAttachments(IOrganizationService service, Guid emailId, string subject, string attachmentBody, string ext)
        {
            Entity attachment = new Entity("activitymimeattachment");
            attachment["subject"] = subject;
            string fileName = $"{subject}.{ext}";
            attachment["filename"] = fileName;
            attachment["body"] = attachmentBody;
            attachment["mimetype"] = MimeMapping.GetMimeMapping(fileName);
            attachment["attachmentnumber"] = 1;
            attachment["objectid"] = new EntityReference("email", emailId);
            attachment["objecttypecode"] = "email";
            service.Create(attachment);
        }
    }
}