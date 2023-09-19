# ExchangeOnlineAdminRestApi

Invoke Commands Exchange Online PowerShell REST API with C#

## Get Started

ExchangeOnlineAdminRestApi can be installed using the Nuget package manager or the dotnet CLI.

```
dotnet add package ExchangeOnlineAdminRestApi
```

[Register the application in Azure AD and Assign API permissions to the application](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#step-1-register-the-application-in-azure-ad)

### Examples

**Acquire AccessToken For Client**

```
using Microsoft.Identity.Client;


var tenantId = "ac6ae0c3-4916-...";
var clientId = "c93766bc-f064-...";
var clientSecret = "7Va8Q~Ua-...";

var app = ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithClientSecret(clientSecret)
    .WithTenantId(tenantId)
    .Build();

var authResult = await app
    .AcquireTokenForClient(scopes: new[] { "https://outlook.office365.com/.default" })
    .ExecuteAsync();

var accessToken = authResult.AccessToken;
```

**Tip AccessToken:** You need to use a cache - either the in-memory one (as per the sample above) or a persisted one.

**Invoke Commands Exchange PowerShell**

```
using ExchangeOnlineAdminRestApi;


var exchangeOnline = new ExchangeOnline(tenantId);

var parameters = new Hashtable();
parameters.Add("Identity", "name-mailbox");
parameters.Add("AccessRights", "SendAs");

var response = await exchangeOnline.InvokeCommand(accessToken, "Get-RecipientPermission", parameters);
```

Convert result on object

```
var response = await exchangeOnline.InvokeCommand<YouClassResult>(accessToken, "Get-RecipientPermission", parameters);
```

### Documentation

- Coming soon.

## About

ExchangeOnlineAdminRestApi was developed by [Claudiney Queiroz](https://claudineyqueiroz.dev).

## License

Copyright © 2023 Claudiney Queiroz.

The project is licensed under the [GNU AGPLv3](https://github.com/claudineyqr/ExchangeOnlineAdminRestApi/blob/master/LICENSE).
