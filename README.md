# ExchangeOnlineAdminRestApi

[![NuGet](https://img.shields.io/nuget/v/ExchangeOnlineAdminRestApi)](https://nuget.org/packages/ExchangeOnlineAdminRestApi) [![Nuget](https://img.shields.io/nuget/dt/ExchangeOnlineAdminRestApi)](https://nuget.org/packages/ExchangeOnlineAdminRestApi) [![Donate](https://img.shields.io/static/v1?label=Paypal&message=Donate&color=informational)](https://www.paypal.com/donate/?hosted_button_id=4CTJ6CTWHGUMA)

Invoke Commands Exchange Online PowerShell REST API with C#

## Get Started

ExchangeOnlineAdminRestApi can be installed using the Nuget package manager or the dotnet CLI.

```
dotnet add package ExchangeOnlineAdminRestApi
```

**[Register the application in Azure AD and Assign API permissions to the application](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#step-1-register-the-application-in-azure-ad)**

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

var response = await exchangeOnline.InvokeCommandAsync(accessToken, "Get-RecipientPermission", parameters);
```

```
using ExchangeOnlineAdminRestApi;


var exchangeOnline = new ExchangeOnline(tenantId);

var parameters = new Hashtable();
parameters.Add("Identity", "name-mailbox");
parameters.Add("AccessRights", "SendAs");

var commandOptions = new CommandOptions
{
    CmdletName = "Get-RecipientPermission",
    Parameters = parameters,
    TimeoutInSeconds = 200
};

var response = await exchangeOnline.InvokeCommandAsync(accessToken, commandOptions);
```

Convert result on object

```
var response = await exchangeOnline.InvokeCommandAsync<YouClassResult>(accessToken, "Get-RecipientPermission", parameters);
```

```
var response = await exchangeOnline.InvokeCommandAsync<YouClassResult>(accessToken, commandOptions);
```

### Documentation

- Coming soon.

## About

ExchangeOnlineAdminRestApi was developed by [Claudiney Queiroz](https://claudineyqueiroz.dev).

## License

Copyright © 2023 Claudiney Queiroz.

The project is licensed under the [GNU AGPLv3](https://github.com/claudineyqr/ExchangeOnlineAdminRestApi/blob/master/LICENSE).
