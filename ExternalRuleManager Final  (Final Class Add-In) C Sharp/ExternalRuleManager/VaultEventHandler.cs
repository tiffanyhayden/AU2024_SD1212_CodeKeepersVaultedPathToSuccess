using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;

namespace ExternalRuleManager
{
    public class VaultEventHandler
    {

        public VaultEventHandler()
        {
            VDF.Vault.Library.ConnectionManager.ConnectionEstablished += OnConnectionEstablished;
            VDF.Vault.Library.ConnectionManager.ConnectionReleased += OnConnectionReleased;
        }

        private void OnConnectionEstablished(object? sender, ConnectionEventArgs e)
        {
            VaultConn.InitializeConnection();
            VaultUtilities.GetLatestFilesByLifecycleState(Globals.ExternalRuleName, "Released");
            Globals.InvAppRibbon.UIEnable(Globals.InvApp.ActiveDocument);
            Globals.InvAppRibbon.LoadRules();



        }

        private void OnConnectionReleased(object? sender, ConnectionEventArgs e)
        {
            Globals.InvAppRibbon.UIDisable(Globals.InvApp.ActiveDocument);
        }



    }
}
