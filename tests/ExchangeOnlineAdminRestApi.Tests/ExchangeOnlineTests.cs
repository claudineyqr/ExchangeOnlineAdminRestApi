// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using Microsoft.Identity.Client;
using System;
using System.Collections;
using System.Threading.Tasks;
using Xunit;

namespace ExchangeOnlineAdminRestApi.Tests;

public class ExchangeOnlineTests
{
    private readonly string _tenantId = Environment.GetEnvironmentVariable("EX_TENANTID");

    [Theory]
    [InlineData("", "")]
    [InlineData("a", "")]
    [InlineData(null, null)]
    [InlineData(null, "")]
    [InlineData("a", null)]
    public void WhenExchangeOnline_GivenContructorWithParamentersNull_ThenReturnArgumentNullException(string tenantId, string apiVersion)
    {
        Assert.Throws<ArgumentNullException>(() => new ExchangeOnline(tenantId, apiVersion));
    }

    [Fact]
    public async Task InvokeCommand_ResultString()
    {
        // Arrange
        var exchangeOnline = new ExchangeOnline(_tenantId);

        var parameters = new Hashtable
        {
            { "Identity", "test01" },
            { "Trustee", "claudiney" },
            { "AccessRights", "SendAs" },
            { "Credencial", "" },
            { "AuthenticationType", "" }
        };

        // Act
        var response = await exchangeOnline.InvokeCommand(await GetAccessToken(), "Get-RecipientPermission", parameters);

        // Assert
        Assert.NotNull(response);
    }

    [Fact]
    public async Task InvokeCommandAsync_ResultString()
    {
        // Arrange
        var exchangeOnline = new ExchangeOnline(_tenantId);

        var parameters = new Hashtable
        {
            { "Identity", "test01" },
            { "Trustee", "claudiney" },
            { "AccessRights", "SendAs" },
            { "Credencial", "" },
            { "AuthenticationType", "" }
        };

        // Act
        var response = await exchangeOnline.InvokeCommandAsync(await GetAccessToken(), "Get-RecipientPermission", parameters);

        // Assert
        Assert.NotNull(response);
    }

    [Fact]
    public async Task InvokeCommand_ResultObject()
    {
        // Arrange
        var exchangeOnline = new ExchangeOnline(_tenantId);

        var parameters = new Hashtable
        {
            { "Identity", "test01" },
            { "Trustee", "claudiney" },
            { "AccessRights", "SendAs" }
        };

        // Act
        var response = await exchangeOnline.InvokeCommand<RecipientPermissionResponse>(await GetAccessToken(), "Get-RecipientPermission", parameters);

        // Assert
        Assert.NotNull(response);
        Assert.IsType<RecipientPermissionResponse>(response);
    }

    [Fact]
    public async Task InvokeCommandAsync_ResultObject()
    {
        // Arrange
        var exchangeOnline = new ExchangeOnline(_tenantId);

        var parameters = new Hashtable
        {
            { "Identity", "test01" },
            { "Trustee", "claudiney" },
            { "AccessRights", "SendAs" }
        };

        // Act
        var response = await exchangeOnline.InvokeCommandAsync<RecipientPermissionResponse>(await GetAccessToken(), "Get-RecipientPermission", parameters);

        // Assert
        Assert.NotNull(response);
        Assert.IsType<RecipientPermissionResponse>(response);
    }

    private async Task<string> GetAccessToken()
    {
        var clientId = Environment.GetEnvironmentVariable("EX_CLIENTID");
        var clientSecret = Environment.GetEnvironmentVariable("EX_CLIENTSECRET");

        var app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithTenantId(_tenantId)
            .Build();

        var authResult = await app
            .AcquireTokenForClient(scopes: new[] { "https://outlook.office365.com/.default" })
            .ExecuteAsync();

        return authResult.AccessToken;
    }
}