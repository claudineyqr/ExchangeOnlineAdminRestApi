// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright � 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using Microsoft.Identity.Client;
using System.Collections;
using System.Threading.Tasks;
using Xunit;

namespace ExchangeOnlineAdminRestApi.Tests
{
    public class ExchangeOnlineTests
    {
        private const string _tenantId = "1225b609-47f2-...";

        [Fact]
        public async Task InvokeCommand_ResultString()
        {
            // Arrange
            var exchangeOnline = new ExchangeOnline(_tenantId);

            var parameters = new Hashtable
            {
                { "Identity", "xxxxxx" },
                { "Trustee", "yyyyy" },
                { "AccessRights", "SendAs" }
            };

            // Act
            var response = await exchangeOnline.InvokeCommand(await GetAccessToken(), "Get-RecipientPermission", parameters);

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
                { "Identity", "xxxxxx" },
                { "Trustee", "yyyyy" },
                { "AccessRights", "SendAs" }
            };

            // Act
            var response = await exchangeOnline.InvokeCommand<RecipientPermissionResponse>(await GetAccessToken(), "Get-RecipientPermission", parameters);

            // Assert
            Assert.NotNull(response);
            Assert.IsType<RecipientPermissionResponse>(response);
        }

        private static async Task<string> GetAccessToken()
        {
            var clientId = "9c42f740-dcff-...";
            var clientSecret = "7Va8Q~Ua-...";

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
}