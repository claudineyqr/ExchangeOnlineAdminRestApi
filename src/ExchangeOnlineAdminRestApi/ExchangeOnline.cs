// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using Newtonsoft.Json;
using System;
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
            _tenantId = string.IsNullOrEmpty(tenantId) ? throw new ArgumentNullException(nameof(tenantId)) : tenantId;
            _apiVersion = string.IsNullOrEmpty(apiVersion) ? throw new ArgumentNullException(nameof(apiVersion)) : apiVersion;
        }

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="cmdletName"></param>
        /// <param name="parameters"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        [Obsolete("InvokeCommand is deprecated, please use InvokeCommandAsync")]
        public async Task<string> InvokeCommand(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default) =>
            await InvokeCommandAsync(accessToken, cmdletName, parameters, cancellationToken);

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="cmdletName"></param>
        /// <param name="parameters"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<string> InvokeCommandAsync(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default)
        {
            var commandOptions = new CommandOptions
            {
                CmdletName = cmdletName,
                Parameters = parameters
            };
            return await InvokeCommandAsync(accessToken, commandOptions, cancellationToken);
        }

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="commandOptions"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        [Obsolete("InvokeCommand is deprecated, please use InvokeCommandAsync")]
        public async Task<string> InvokeCommand(string accessToken, CommandOptions commandOptions, CancellationToken cancellationToken = default) =>
            await InvokeCommandAsync(accessToken, commandOptions, cancellationToken);

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="commandOptions"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<string> InvokeCommandAsync(string accessToken, CommandOptions commandOptions, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(accessToken)) throw new ArgumentNullException(nameof(accessToken));
            if (commandOptions is null) throw new ArgumentNullException(nameof(commandOptions));
            if (string.IsNullOrEmpty(commandOptions.CmdletName)) throw new ArgumentNullException(nameof(commandOptions.CmdletName));

            var body = new
            {
                CmdletInput = new
                {
                    commandOptions.CmdletName,
                    Parameters = ClearParameters(commandOptions.Parameters) ?? new Hashtable()
                }
            };

            var json = JsonConvert.SerializeObject(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var client = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(commandOptions.TimeoutInSeconds)
            };
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
        [Obsolete("InvokeCommand is deprecated, please use InvokeCommandAsync")]
        public async Task<TResult> InvokeCommand<TResult>(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default) =>
            await InvokeCommandAsync<TResult>(accessToken, cmdletName, parameters, cancellationToken);

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="accessToken"></param>
        /// <param name="cmdletName"></param>
        /// <param name="parameters"></param> 
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TResult> InvokeCommandAsync<TResult>(string accessToken, string cmdletName, Hashtable parameters = null, CancellationToken cancellationToken = default)
        {
            var commandOptions = new CommandOptions
            {
                CmdletName = cmdletName,
                Parameters = parameters
            };
            return await InvokeCommandAsync<TResult>(accessToken, commandOptions, cancellationToken);
        }

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="accessToken"></param>
        /// <param name="commandOptions"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        [Obsolete("InvokeCommand is deprecated, please use InvokeCommandAsync")]
        public async Task<TResult> InvokeCommand<TResult>(string accessToken, CommandOptions commandOptions, CancellationToken cancellationToken = default) =>
            await InvokeCommandAsync<TResult>(accessToken, commandOptions, cancellationToken);

        /// <summary>
        /// Invoke Commands Exchange Online PowerShell REST API
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="accessToken"></param>
        /// <param name="commandOptions"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<TResult> InvokeCommandAsync<TResult>(string accessToken, CommandOptions commandOptions, CancellationToken cancellationToken = default)
        {
            var response = await InvokeCommandAsync(accessToken, commandOptions, cancellationToken);
            return JsonConvert.DeserializeObject<TResult>(response);
        }

        private string GetUriAdminApi()
        {
            return string.Format(URL_INVOKE_COMMAND, _apiVersion, _tenantId);
        }

        private static Hashtable ClearParameters(Hashtable parameters)
        {
            parameters?.Remove("Credencial");
            parameters?.Remove("AuthenticationType");
            return parameters;
        }
    }
}