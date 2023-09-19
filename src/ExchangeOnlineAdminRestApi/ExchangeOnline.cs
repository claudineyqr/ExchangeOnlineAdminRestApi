﻿// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using Newtonsoft.Json;
using System.Collections;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExchangeOnlineAdminRestApi
{
    public class ExchangeOnline
    {
        private const string URL_INVOKE_COMMAND = "https://outlook.office365.com/AdminApi/{0}/{1}/InvokeCommand";
        private readonly string _tenantId;
        private readonly string _apiVersion;

        public ExchangeOnline(string tenantId, string apiVersion = "beta")
        {
            _tenantId = tenantId;
            _apiVersion = apiVersion;
        }

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="cmdletName"></param>
        /// <param name="parameters"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<string> InvokeCommand(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default)
        {
            var body = new
            {
                CmdletInput = new
                {
                    CmdletName = cmdletName,
                    Parameters = parameters ?? new Hashtable()
                }
            };

            var json = JsonConvert.SerializeObject(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var response = await client.PostAsync(GetUriAdminApi(), content, cancellationToken);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
                return responseBody;

            throw new ExchangeOnlineException(responseBody, response.StatusCode);
        }

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="accessToken"></param>
        /// <param name="cmdletName"></param>
        /// <param name="parameters"></param> 
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TResult> InvokeCommand<TResult>(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default)
        {
            var response = await InvokeCommand(accessToken, cmdletName, parameters, cancellationToken);
            return JsonConvert.DeserializeObject<TResult>(response);
        }

        private string GetUriAdminApi()
        {
            return string.Format(URL_INVOKE_COMMAND, _apiVersion, _tenantId);
        }
    }
}