using ExternalRuleManager.Properties;
using Inventor;
using System.Diagnostics;
using ACW = Autodesk.Connectivity.WebServices;
using File = System.IO.File;
using VDF = Autodesk.DataManagement.Client.Framework;

namespace ExternalRuleManager
{
    /// <summary>
    /// Manages the creation and customization of a custom ribbon in the Autodesk Inventor UI.
    /// </summary>
    public class CustomRibbon
    {
        #region Properties

        public static ComboBoxDefinition ExternalRules { get; set; }
        public ComboBoxDefinition CurrentLifecycleState { get; set; }



        public ButtonDefinition GetLatestAllRules { get; set; }
        public ButtonDefinition GetLatestReleasedAllRules { get; set; }
        public ButtonDefinition CheckIn { get; set; }
        public ButtonDefinition CheckOut { get; set; }
        public ButtonDefinition UndoCheckOut { get; set; }
        public ButtonDefinition MakeLocalCopy { get; set; }
        public ButtonDefinition OverwriteRuleOnDiskWithCopy { get; set; }




        private bool isExternalRulesInitialized = false;

        private RibbonTab externalRuleManagerRibbonTab;

        private RibbonPanel manageRibbonPanel;
        private RibbonPanel statusRibbonPanel;
        private RibbonPanel allRulesRibbonPanel;
        private RibbonPanel lifecycleRibbonPanel;
        private RibbonPanel LocalRibbonPanel;



        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomRibbon"/> class.
        /// Creates the TestButton and adds it to the specified ribbons in the Inventor UI.
        /// </summary>
        public CustomRibbon()
        {
            ExternalRules = Utilities.CreateComboBoxDef("External Rules", "ExternalRules", CommandTypesEnum.kQueryOnlyCmdType, 200);
            AddExternalRules();
            ExternalRules.OnSelect += ExternalRules_OnSelect;

            Ribbon partRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            Ribbon assemblyRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            Ribbon drawingRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            Ribbon zeroDocRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["ZeroDoc"];

            AddToRibbon(partRibbon);
            AddToRibbon(assemblyRibbon);
            AddToRibbon(drawingRibbon);
            AddToRibbon(zeroDocRibbon);


            isExternalRulesInitialized = true;
        }

        #endregion

        /// <summary>
        /// Converts a byte array to an Image.
        /// </summary>
        /// <param name="byteArray">The byte array to convert.</param>
        /// <returns>The converted Image object.</returns>
        private Image ByteArrayToImage(byte[] byteArray)
        {
            using (MemoryStream ms = new MemoryStream(byteArray))
            {
                return Image.FromStream(ms);
            }
        }

        /// <summary>
        /// Adds the TestButton to a specific ribbon in the Inventor UI.
        /// </summary>
        /// <param name="targetRibbon">The ribbon to which the TestButton will be added.</param>
        private void AddToRibbon(Ribbon targetRibbon)
        {
            externalRuleManagerRibbonTab = targetRibbon.RibbonTabs.Add("External Rule Manager", "id_ExternalRuleManager" + targetRibbon.InternalName, Globals.InvAppGuidID);

            manageRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Manage", "id_Manage" + targetRibbon.InternalName, Globals.InvAppGuidID);

            allRulesRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("All Rules", "id_AllRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            statusRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Status", "id_Status" + targetRibbon.InternalName, Globals.InvAppGuidID);

            lifecycleRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("LifeCycle", "id_Lifecycle" + targetRibbon.InternalName, Globals.InvAppGuidID);

            LocalRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Local (C:)", "id_Local" + targetRibbon.InternalName, Globals.InvAppGuidID);

            manageRibbonPanel.CommandControls.AddComboBox(ExternalRules);

            if (targetRibbon.InternalName != "ZeroDoc")
            {
                if (targetRibbon.InternalName == "Part")
                {
                    // Additional logic for Part ribbon can be added here
                }
                else if (targetRibbon.InternalName == "Assembly")
                {
                    // Additional logic for Assembly ribbon can be added here
                }
                else if (targetRibbon.InternalName == "Drawing")
                {
                    // Additional logic for Drawing ribbon can be added here
                }
            }
        }

        #region Ui Events



        private void ExternalRules_OnSelect(NameValueMap context)
        {
            if(isExternalRulesInitialized)
            {
                CreateAndLoadOtherButtons();
            }
            
        }

        private void GetLatestAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestOnFolder(Globals.ExternalRuleName);

        }

        private void GetLatestReleasedAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestFilesByLifecycleState(Globals.ExternalRuleName, "");
        }

        private void UndoCheckOut_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex]; 
            if(ExternalRules.ListIndex != 1)
            {
                VaultFileUtilities.File_UndoCheckOut(selectedItemName);
            }
            
        }

        private void CheckOut_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                VaultFileUtilities.File_CheckOut(file);
            }

        }

        private void CheckIn_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                VaultFileUtilities.File_CheckIn(selectedItemName, "");
            }

        }

        private void CurrentLifecycleState_OnSelect(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {

                VaultLifecycleUtilities.GetLifecycleState(selectedItemName);
            }

        }

        private void MakeLocalCopy_OnExecute(NameValueMap context)
        {
            FileUtilities.MakeLocalCopy();
        }


        private void OverwriteRuleOnDiskWithCopy_OnExecute(NameValueMap context)
        {
            FileUtilities.OverwriteRuleOnDiskWithCopy();
        }



        #endregion


        #region Methods

        public void AddExternalRules()
        {

            VaultConn.InitializeConnection();
            ACW.File[]? files = VaultUtilities.GetFilesByFolder(Globals.ExternalRuleName);



            if (files != null && files.Length > 0)
            {
                ExternalRules.Clear();

                List<ACW.File> sortedFiles = files.OrderBy(file => file.Name).ToList();

                ExternalRules.AddItem("No External Rule Selected");

                foreach (ACW.File file in sortedFiles)
                {
                    ExternalRules.AddItem(file.Name);
                }


                ExternalRules.ListIndex = 1;

            }

        }


        public void CreateAndLoadOtherButtons()
        {

            
            Image image16x16 = ByteArrayToImage(Resources._16x16);
            Image image32x32 = ByteArrayToImage(Resources._32x32);
            Image checkIn16x16 = ByteArrayToImage(Resources.checkin_16x16);
            Image checkIn32x32 = ByteArrayToImage(Resources.checkin_32x32);



            GetLatestAllRules = Utilities.CreateButtonDef("Get Latest", "GetLatestAllRules", "", image16x16, image32x32);

            if(GetLatestAllRules != null)
            {
                allRulesRibbonPanel.CommandControls.AddButton(GetLatestAllRules);
                GetLatestAllRules.OnExecute += GetLatestAllRules_OnExecute;
            }


            GetLatestReleasedAllRules = Utilities.CreateButtonDef("Get Latest Released Version", "GetLatestReleasedVersionAllRules", "", image16x16, image32x32);

            if (GetLatestReleasedAllRules != null)
            {
                allRulesRibbonPanel.CommandControls.AddButton(GetLatestReleasedAllRules);
                GetLatestReleasedAllRules.OnExecute += GetLatestReleasedAllRules_OnExecute;
            }

            CheckIn = Utilities.CreateButtonDef("", "Checkin", "", checkIn16x16, checkIn32x32);

            if (CheckIn != null)
            {
                statusRibbonPanel.CommandControls.AddButton(CheckIn, true);
                CheckIn.OnExecute += CheckIn_OnExecute;
            }

            CheckOut = Utilities.CreateButtonDef("Check Out", "Checkout", "", image16x16, image32x32);

            if (CheckOut != null)
            {
                statusRibbonPanel.CommandControls.AddButton(CheckOut, true);
                CheckOut.OnExecute += CheckOut_OnExecute;
            }


            UndoCheckOut = Utilities.CreateButtonDef("Undo Check Out", "UndoCheckOut", "", image16x16, image16x16);

            if (UndoCheckOut != null)
            {
                statusRibbonPanel.CommandControls.AddButton(UndoCheckOut, true);
                UndoCheckOut.OnExecute += UndoCheckOut_OnExecute;
            }


            CurrentLifecycleState = Utilities.CreateComboBoxDef("Current Lifecycle State", "CurrentLifecycleState", CommandTypesEnum.kQueryOnlyCmdType, 200);

            if (CurrentLifecycleState != null)
            {
                lifecycleRibbonPanel.CommandControls.AddComboBox(CurrentLifecycleState);
                string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
                if (ExternalRules.ListIndex != 1)
                {
                    CurrentLifecycleState.AddItem(VaultLifecycleUtilities.GetLifecycleState(selectedItemName));
                    CurrentLifecycleState.ListIndex = 1;
                }
            }

            MakeLocalCopy = Utilities.CreateButtonDef("Make Local Copy", "MakeLocalCopy", "", image16x16, image32x32);

            if (MakeLocalCopy != null)
            {
                LocalRibbonPanel.CommandControls.AddButton(MakeLocalCopy, true);
                MakeLocalCopy.OnExecute += MakeLocalCopy_OnExecute;
            }


            OverwriteRuleOnDiskWithCopy = Utilities.CreateButtonDef("Overwrite Rule\r\nWith Copy", "OverwriteRuleOnDiskWithCopy", "", image16x16, image32x32);

            if (OverwriteRuleOnDiskWithCopy != null)
            {
                LocalRibbonPanel.CommandControls.AddButton(OverwriteRuleOnDiskWithCopy, true);
                OverwriteRuleOnDiskWithCopy.OnExecute += OverwriteRuleOnDiskWithCopy_OnExecute;
            }



        }




        #endregion


    }
}