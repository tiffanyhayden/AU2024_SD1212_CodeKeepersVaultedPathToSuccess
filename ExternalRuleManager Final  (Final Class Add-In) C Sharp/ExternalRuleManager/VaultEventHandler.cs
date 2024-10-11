using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;

namespace ExternalRuleManager
{
    /// <summary>
    /// Handles events related to Autodesk Vault connection establishment and release.
    /// </summary>
    public class VaultEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="VaultEventHandler"/> class and subscribes to Vault connection events.
        /// </summary>
        public VaultEventHandler()
        {
            // Subscribing to Vault connection established and released events
            VDF.Vault.Library.ConnectionManager.ConnectionEstablished += OnConnectionEstablished;
            VDF.Vault.Library.ConnectionManager.ConnectionReleased += OnConnectionReleased;
        }

        /// <summary>
        /// Event handler that is called when a connection to Vault is established.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">Event data containing information about the established connection.</param>
        /// <remarks>
        /// This method initializes the Vault connection, retrieves the latest files by lifecycle state, and updates the Inventor ribbon UI.
        /// </remarks>
        private void OnConnectionEstablished(object? sender, ConnectionEventArgs e)
        {
            // Initialize the Vault connection
            VaultConn.InitializeConnection();

            // Retrieve the latest files by lifecycle state "Released"
            VaultUtilities.GetLatestFilesByLifecycleState(Globals.ExternalRuleName, "Released");

            // Enable and refresh the Inventor ribbon UI
            Globals.InvAppRibbon.UIEnable(Globals.InvApp.ActiveDocument);
            Globals.InvAppRibbon.LoadExternalRules();
        }

        /// <summary>
        /// Event handler that is called when the connection to Vault is released.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">Event data containing information about the released connection.</param>
        /// <remarks>
        /// This method disables the UI elements of the Inventor ribbon when the Vault connection is released.
        /// </remarks>
        private void OnConnectionReleased(object? sender, ConnectionEventArgs e)
        {
            // Disable the Inventor ribbon UI when the Vault connection is released
            Globals.InvAppRibbon.UIDisable(Globals.InvApp.ActiveDocument);
        }
    }
}
