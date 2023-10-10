// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using System.Collections;

namespace ExchangeOnlineAdminRestApi;

public class CommandOptions
{
    public string CmdletName { get; set; }

    /// <summary>
    /// Parameters that are available on the cmdlet and values that each parameter accepts.
    /// </summary>
    public Hashtable Parameters { get; set; } = new Hashtable();

    /// <summary>
    /// Sets the timespan to wait before the request times out. Default: 100
    /// </summary>
    public int TimeoutInSeconds { get; set; } = 100;
}
