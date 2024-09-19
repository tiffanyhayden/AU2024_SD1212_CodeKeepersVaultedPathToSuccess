using Inventor;
using System.Runtime.InteropServices;
using static ExternalRuleManager.Globals;

namespace ExternalRuleManager
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [ProgId("TestAddin2.StandardAddInServer")]
    [Guid(InvAppGuid)]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;

        //user interface event
        private UserInterfaceEvents m_uiEvents;



        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            Globals.InvApp = addInSiteObject.Application;
            Globals.InvAppEvents = new InventorEventHandler();
            Globals.VaultAppEvents = new VaultEventHandler();
            //Globals.InvAppRibbon = new CustomRibbon();



            m_uiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;

        }
        public void Deactivate()
        {
            m_uiEvents = null;
            Globals.InvApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        #endregion

    }
}
