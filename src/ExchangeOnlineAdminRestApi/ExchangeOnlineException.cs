// 
// This file is part of Project ExchangeOnlineAdminRestApi
// Author: Claudiney Queiroz (claudiney.queiroz@gmail.com, https://www.linkedin.com/in/claudineyqr/)
// ProjectUrl: https://github.com/claudineyqr/ExchangeOnlineAdminRestApi
// Copyright © 2023 Claudiney Queiroz.
// License: GNU Affero General Public License v3.0 (GNU AGPLv3)

using System;
using System.Net;

namespace ExchangeOnlineAdminRestApi
{
    public class ExchangeOnlineException : Exception
    {
        public ExchangeOnlineException(string message, HttpStatusCode statusCode) :
            base(message, new Exception($"HttpStatusCode: {statusCode}"))
        {

        }
    }
}
