using VDF = Autodesk.DataManagement.Client.Framework;
using VltBase = Connectivity.Application.VaultBase;

namespace ExternalRuleManager
{
    /// <summary>
    /// Provides methods and properties to manage and access the connection to Autodesk Vault.
    /// </summary>
    public static class VaultConn
    {
        /// <summary>
        /// The active Vault connection. This is <c>null</c> if no connection is currently established.
        /// </summary>
        private static VDF.Vault.Currency.Connections.Connection? _ActiveConnection;

        /// <summary>
        /// Gets or sets the active Vault connection. If no connection is active, it attempts to retrieve the connection from the <see cref="VltBase.ConnectionManager"/>.
        /// </summary>
        /// <returns>
        /// The current <see cref="VDF.Vault.Currency.Connections.Connection"/> object if a connection is active, or the connection from the Vault connection manager if available.
        /// </returns>
        public static VDF.Vault.Currency.Connections.Connection? ActiveConnection
        {
            get
            {
                if (_ActiveConnection == null)
                {
                    // Return the connection from the Vault connection manager if available.
                    return VltBase.ConnectionManager.Instance.Connection;
                }
                else
                {
                    return _ActiveConnection;
                }
            }
            set
            {
                _ActiveConnection = value;
            }
        }

        /// <summary>
        /// Checks whether there is an active connection to Autodesk Vault.
        /// </summary>
        /// <returns>
        /// <c>true</c> if there is an active Vault connection; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsConnected()
        {
            return _ActiveConnection != null;
        }

        /// <summary>
        /// Initializes a connection to Autodesk Vault by obtaining the connection from the <see cref="VltBase.ConnectionManager"/>.
        /// </summary>
        /// <returns>
        /// The initialized <see cref="VDF.Vault.Currency.Connections.Connection"/> object, or <c>null</c> if the connection could not be established.
        /// </returns>
        /// <remarks>
        /// If the connection cannot be obtained, the <see cref="ActiveConnection"/> property will remain <c>null</c>. 
        /// This method does not throw exceptions but can be enhanced with additional error handling if needed.
        /// </remarks>
        public static VDF.Vault.Currency.Connections.Connection InitializeConnection()
        {
            var connection = VltBase.ConnectionManager.Instance.Connection;

            if (connection != null)
            {
                ActiveConnection = connection;
            }
            else
            {
                // Optionally handle the case where the connection could not be obtained.
            }

            return ActiveConnection;
        }
    }
}
