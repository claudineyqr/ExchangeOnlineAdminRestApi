// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using Newtonsoft.Json;
using System.Collections.Generic;

namespace ExchangeOnlineAdminRestApi.Tests;

internal class RecipientPermissionResponse
{
    [JsonProperty("@odata.context")]
    public string OdataContext { get; set; }

    [JsonProperty("adminapi.warnings@odata.type")]
    public string AdminapiWarningsOdataType { get; set; }

    [JsonProperty("@adminapi.warnings")]
    public List<object> AdminapiWarnings { get; set; }

    [JsonProperty("value")]
    public List<Value> Value { get; set; }
}

internal class Value
{
    public string Identity { get; set; }
    public string Trustee { get; set; }
    public string AccessControlType { get; set; }
    public string AccessRightsOdataType { get; set; }
    public List<string> AccessRights { get; set; }
    public bool IsInherited { get; set; }
    public string InheritanceType { get; set; }
    public string TrusteeSidString { get; set; }
    public bool IsValid { get; set; }
}
