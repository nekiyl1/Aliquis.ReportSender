# Aliquis.ReportSender

Решение для отправки отчётов Reporting Services по электронной почте для Microsoft Dynamics CRM 2016/365 on-premises. После установки решения его можно использовать для автоматической отправки отчётов на электронную почту пользователя, например, добавив приложение с вызовом действия al_AliquisReportSender в планировщик.

### Примеры использования:

**C#**
```
var request = new OrganizationRequest("al_AliquisReportSender")
{
	["toUserId"] = "28BD6252-D22B-42E6-9DED-C1DB35455DEE",
	["subject"] = "Отчёт",
	["parameters"] = "{'test':'test'}",
	["format"] = "EXCEL",
	["userName"] = "userName",
	["password"] = "password",
	["domain"] = "domain",
	["url"] = "https://crmurl/orgname/",
	["reportid"] = "84FE8E4F-49C6-4DA1-BAC5-E9BE1D690C0A",
};
var response = service.Execute(request);
```

**JavaScript**
```
var parameters = {};
parameters.reportid = "84FE8E4F-49C6-4DA1-BAC5-E9BE1D690C0A";
parameters.parameters = "{'test':'test'}";
parameters.format = "EXCEL";
parameters.userName = "userName";
parameters.password = "password";
parameters.domain = "domain";
parameters.url = "https://crmurl/orgname/";
parameters.toUserId = "28BD6252-D22B-42E6-9DED-C1DB35455DEE";
parameters.subject = "Отчёт";
var req = new XMLHttpRequest();
req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/al_AliquisReportSender", true);
req.setRequestHeader("OData-MaxVersion", "4.0");
req.setRequestHeader("OData-Version", "4.0");
req.setRequestHeader("Accept", "application/json");
req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
req.onreadystatechange = function() {
    if (this.readyState === 4) {
        req.onreadystatechange = null;
        if (this.status === 204) {
            //Success - No Return Data - Do Something
        } else {
            Xrm.Utility.alertDialog(this.statusText);
        }
    }
};
req.send(JSON.stringify(parameters));
```

**Параметры:**
* reportid – идентификатор отчёта
* parameters – параметры отчёта в формате JSON
* format – формат в котором будет выгружен отчёт (WORD, EXCEL, PDF)
* userName – логин пользователя с доступом в систему
* password – пароль пользователя с доступом в систему
* domain – домен пользователя с доступом в систему
* url – адрес системы в формате https://crmurl/orgname/
* toUserId – идентификатор пользователя которому будет отправлен сформированный отчёт
* subject – тема письма

Скачать решение: https://aliquis.ru/programmirovanie/aliquis-reportsender/