using Inventor;
using static ExternalRuleManager.Globals;

namespace ExternalRuleManager
{
    /// <summary>
    /// Handles Inventor application events such as opening, closing, saving, and activating documents.
    /// </summary>
    public class InventorEventHandler
    {
        /// <summary>
        /// Gets the Inventor application's <see cref="ApplicationEvents"/> object that provides access to document-related events.
        /// </summary>
        protected ApplicationEvents ApplicationEvents { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="InventorEventHandler"/> class and subscribes to relevant Inventor document events.
        /// </summary>
        public InventorEventHandler()
        {
            ApplicationEvents = InvApp.ApplicationEvents;

            // Subscribing to Inventor application events
            ApplicationEvents.OnOpenDocument += OnOpenDocument;
            ApplicationEvents.OnInitializeDocument += OnInitializeDocument;
            ApplicationEvents.OnSaveDocument += OnSaveDocument;
            ApplicationEvents.OnCloseDocument += OnCloseDocument;
            ApplicationEvents.OnActivateDocument += OnActivateDocument;
        }

        /// <summary>
        /// Event handler for when a document is opened in Inventor.
        /// </summary>
        /// <param name="DocumentObject">The opened document object.</param>
        /// <param name="FullDocumentName">The full path of the opened document.</param>
        /// <param name="BeforeOrAfter">Indicates whether the event is fired before or after the document is opened.</param>
        /// <param name="Context">Context information for the event.</param>
        /// <param name="HandlingCode">Outputs the handling code to indicate how the event was handled.</param>
        private void OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            // Default event handling code
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        /// <summary>
        /// Event handler for when a document is initialized in Inventor.
        /// </summary>
        /// <param name="DocumentObject">The initialized document object.</param>
        /// <param name="FullDocumentName">The full path of the initialized document.</param>
        /// <param name="BeforeOrAfter">Indicates whether the event is fired before or after the document is initialized.</param>
        /// <param name="Context">Context information for the event.</param>
        /// <param name="HandlingCode">Outputs the handling code to indicate how the event was handled.</param>
        private void OnInitializeDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            // Default event handling code
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        /// <summary>
        /// Event handler for when a document is saved in Inventor.
        /// </summary>
        /// <param name="DocumentObject">The saved document object.</param>
        /// <param name="BeforeOrAfter">Indicates whether the event is fired before or after the document is saved.</param>
        /// <param name="Context">Context information for the event.</param>
        /// <param name="HandlingCode">Outputs the handling code to indicate how the event was handled.</param>
        private void OnSaveDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            // Default event handling code
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        /// <summary>
        /// Event handler for when a document is closed in Inventor.
        /// </summary>
        /// <param name="DocumentObject">The closed document object.</param>
        /// <param name="FullDocumentName">The full path of the closed document.</param>
        /// <param name="BeforeOrAfter">Indicates whether the event is fired before or after the document is closed.</param>
        /// <param name="Context">Context information for the event.</param>
        /// <param name="HandlingCode">Outputs the handling code to indicate how the event was handled.</param>
        private void OnCloseDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            // Perform custom UI updates after the document is closed
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                Globals.InvAppRibbon.UIEnable(DocumentObject);
                Globals.InvAppRibbon.RefreshUI();
            }

            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        /// <summary>
        /// Event handler for when a document is activated in Inventor.
        /// </summary>
        /// <param name="DocumentObject">The activated document object.</param>
        /// <param name="BeforeOrAfter">Indicates whether the event is fired before or after the document is activated.</param>
        /// <param name="Context">Context information for the event.</param>
        /// <param name="HandlingCode">Outputs the handling code to indicate how the event was handled.</param>
        private void OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            // Perform custom UI updates after the document is activated
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                Globals.InvAppRibbon.UIEnable(DocumentObject);
                Globals.InvAppRibbon.RefreshUI();
            }

            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
    }
}
