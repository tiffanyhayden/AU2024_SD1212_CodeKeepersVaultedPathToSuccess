using Autodesk.Connectivity.WebServices;
using Inventor;
using static ExternalRuleManager.Globals;


namespace ExternalRuleManager
{
    public class InventorEventHandler
    {
        protected ApplicationEvents ApplicationEvents { get; }


        public InventorEventHandler()
        {
            ApplicationEvents = InvApp.ApplicationEvents;

            ApplicationEvents.OnOpenDocument += OnOpenDocument;
            ApplicationEvents.OnInitializeDocument += OnInitializeDocument;
            ApplicationEvents.OnSaveDocument += OnSaveDocument;
            ApplicationEvents.OnCloseDocument += OnCloseDocument;
            ApplicationEvents.OnActivateDocument += OnActivateDocument;

               
        }

        private void OnOpenDocument(_Document DocumentObjects, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {

            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        private void OnInitializeDocument(_Document DocumentObjects, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        private void OnSaveDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        private void OnCloseDocument(_Document DocumentObjects, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }

        private void OnActivateDocument(_Document DocumentObject, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if(BeforeOrAfter == EventTimingEnum.kAfter)
            {
                Globals.InvAppRibbon.UIEnable(DocumentObject);
            }

            HandlingCode = HandlingCodeEnum.kEventNotHandled;


        }


    }
}
