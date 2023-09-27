using System.Collections;

namespace ExchangeOnlineAdminRestApi
{
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
}
