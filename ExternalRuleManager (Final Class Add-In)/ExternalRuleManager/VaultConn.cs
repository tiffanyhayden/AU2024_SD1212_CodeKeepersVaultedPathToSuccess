
using VDF = Autodesk.DataManagement.Client.Framework;
using VltBase = Connectivity.Application.VaultBase;

namespace ExternalRuleManager
{
    public class VaultConn
    {

        private static VDF.Vault.Currency.Connections.Connection? _ActiveConnection;




        public static VDF.Vault.Currency.Connections.Connection? ActiveConnection
        {
            get
            {
                if (_ActiveConnection == null)

                {
                    return VltBase.ConnectionManager.Instance.Connection;
                    //throw new InvalidOperationException("Vault is not connected.");
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

        public static bool IsConnected()
        {
            return _ActiveConnection != null;
        }


        public static VDF.Vault.Currency.Connections.Connection InitializeConnection()
        {
            var connection = VltBase.ConnectionManager.Instance.Connection;

            if (connection != null)
            {
                ActiveConnection = connection;
            }
            else
            {
                // throw new InvalidOperationException("Failed to obtain a Vault connection.");
            }

            return ActiveConnection;

        }

    }

}

