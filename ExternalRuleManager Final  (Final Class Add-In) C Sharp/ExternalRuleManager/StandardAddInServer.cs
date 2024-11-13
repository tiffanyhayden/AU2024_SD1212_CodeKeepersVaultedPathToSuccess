using Inventor;
using System.Runtime.InteropServices;
using static ExternalRuleManager.Globals;

namespace ExternalRuleManager
{
    /// <summary>
    /// The primary AddIn server class that implements the <see cref="ApplicationAddInServer"/> interface.
    /// This interface is required by all Inventor AddIns to facilitate communication between Inventor and the AddIn.
    /// </summary>
    [ProgId("TestAddin2.StandardAddInServer")]
    [Guid(InvAppGuid)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        // Inventor application object.
        private Inventor.Application m_inventorApplication;

        // User interface events object.
        private UserInterfaceEvents m_uiEvents;

        /// <summary>
        /// Initializes a new instance of the <see cref="StandardAddInServer"/> class.
        /// </summary>
        public StandardAddInServer()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        #region ApplicationAddInServer Members

        /// <summary>
        /// Activates the AddIn and initializes the Inventor environment.
        /// This method is called by Inventor when the AddIn is loaded.
        /// </summary>
        /// <param name="addInSiteObject">The <see cref="ApplicationAddInSite"/> object that provides access to the Inventor application.</param>
        /// <param name="firstTime">Indicates whether the AddIn is being loaded for the first time.</param>
        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // Initialize AddIn members.
            
            Globals.InvApp = addInSiteObject.Application;
            Globals.InvAppEvents = new InventorEventHandler();
            Globals.VaultAppEvents = new VaultEventHandler();
            Globals.InvAppRibbon = new CustomRibbon();

            // Set up user interface events
            m_uiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;
        }

        /// <summary>
        /// Deactivates the AddIn and cleans up resources.
        /// This method is called by Inventor when the AddIn is unloaded.
        /// </summary>
        public void Deactivate()
        {
            // Clean up UI events and other resources
            m_uiEvents = null;
            Globals.InvApp = null;

            // Force garbage collection to clean up any remaining objects.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }


        /// <summary>
        /// Exposes an API for the AddIn that can be accessed by external programs.
        /// </summary>
        /// <remarks>
        /// Typically, this would return an object that implements the AddIn's API, allowing external programs to interact with the AddIn.
        /// </remarks>
        public object Automation
        {
            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        void ApplicationAddInServer.ExecuteCommand(int CommandID)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
