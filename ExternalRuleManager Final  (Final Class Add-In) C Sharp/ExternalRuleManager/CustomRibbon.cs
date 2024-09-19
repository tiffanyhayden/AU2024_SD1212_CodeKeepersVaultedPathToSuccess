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
        public ButtonDefinition Get { get; set; }
        public ButtonDefinition CheckIn { get; set; }
        public ButtonDefinition CheckOut { get; set; }
        public ButtonDefinition UndoCheckOut { get; set; }
        public ButtonDefinition MakeLocalCopy { get; set; }
        public ButtonDefinition OverwriteRuleOnDiskWithCopy { get; set; }
        public static ButtonDefinition Refresh { get; set; }



        public bool isExternalRulesInitialized = false;

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

            Image refresh16x16 = ByteArrayToImage(Resources.Refresh_16x16);
            Image refresh32x32 = ByteArrayToImage(Resources.Refresh_32x32);
            Refresh = Utilities.CreateButtonDef("Refresh", "CustomRefresh", "", refresh16x16, refresh32x32);



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
        public   void AddToRibbon(Ribbon targetRibbon)
        {
            externalRuleManagerRibbonTab = targetRibbon.RibbonTabs.Add("External Rule Manager", "id_ExternalRuleManager" + targetRibbon.InternalName, Globals.InvAppGuidID);

            manageRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Manage", "id_Manage" + targetRibbon.InternalName, Globals.InvAppGuidID);

            allRulesRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("All Rules", "id_AllRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            statusRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Status", "id_Status" + targetRibbon.InternalName, Globals.InvAppGuidID);

            lifecycleRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Lifecycle", "id_Lifecycle" + targetRibbon.InternalName, Globals.InvAppGuidID);

            LocalRibbonPanel = externalRuleManagerRibbonTab.RibbonPanels.Add("Local (C:)", "id_Local" + targetRibbon.InternalName, Globals.InvAppGuidID);

            manageRibbonPanel.CommandControls.AddComboBox(ExternalRules);




            //if (targetRibbon.InternalName != "ZeroDoc")
            //{
            //    if (targetRibbon.InternalName == "Part")
            //    {
            //        // Additional logic for Part ribbon can be added here
            //    }
            //    else if (targetRibbon.InternalName == "Assembly")
            //    {
            //        // Additional logic for Assembly ribbon can be added here
            //    }
            //    else if (targetRibbon.InternalName == "Drawing")
            //    {
            //        // Additional logic for Drawing ribbon can be added here
            //    }
            //}
        }




        #region Ui Events

        private void Refresh_OnExecute(NameValueMap context)
        {
            if (isExternalRulesInitialized)
            {
                RefreshUI();
            }

        }

        private void ExternalRules_OnSelect(NameValueMap context)
        {
            if(isExternalRulesInitialized)
            {
                CreateAndLoadOtherButtons();
                
                RefreshUI();
                
                
            }
            
        }

        private void GetLatestAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestOnFolder(Globals.ExternalRuleName);

            RefreshUI();

        }

        private void GetLatestReleasedAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestFilesByLifecycleState(Globals.ExternalRuleName, "Released");

            RefreshUI();
        }

        private void UndoCheckOut_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex]; 
            if(ExternalRules.ListIndex != 1)
            {
                VaultFileUtilities.File_UndoCheckOut(selectedItemName);
            }

            RefreshUI();

        }

        private void Get_OnExecute(NameValueMap context)
        {
            

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                VaultFileUtilities.File_GetLatest(selectedItemName);
            }

            RefreshUI();

        }

        private void CheckOut_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                VaultFileUtilities.File_CheckOut(file);
            }

            RefreshUI();

        }

        private void CheckIn_OnExecute(NameValueMap context)
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            if (ExternalRules.ListIndex != 1)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                VaultFileUtilities.File_CheckIn(selectedItemName, "");
            }

            RefreshUI();

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

            RefreshUI();
        }


        private void OverwriteRuleOnDiskWithCopy_OnExecute(NameValueMap context)
        {
            FileUtilities.OverwriteRuleOnDiskWithCopy();

            RefreshUI();
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


        public void RefreshUI()
        {

            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            bool isCheckedOut = false;

            if (ExternalRules.ListIndex == 1)
            {
                Get.Enabled = false;
                CheckIn.Enabled = false;
                CheckOut.Enabled = false;
                MakeLocalCopy.Enabled = false;
                OverwriteRuleOnDiskWithCopy.Enabled = false;
                UndoCheckOut.Enabled = false;
                CurrentLifecycleState.Clear();
                CurrentLifecycleState.ListIndex = 0;

            }
            else
            {
                if (ExternalRules.ListIndex != 1)
                {
                    ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                    if (file != null)
                    {
                        isCheckedOut = VaultFileUtilities.File_IsCheckedOut(file.Name);

                        Get.Enabled = true;
                        CheckOut.Enabled = !isCheckedOut;
                        UndoCheckOut.Enabled = isCheckedOut;
                        CheckIn.Enabled = isCheckedOut;
                        MakeLocalCopy.Enabled = true;
                        OverwriteRuleOnDiskWithCopy.Enabled = true;


                        ResetCurrentLifecycleDropdown();


                    }
                }
            }
        }



        public void CreateAndLoadOtherButtons()
        {

            
            Image image16x16 = ByteArrayToImage(Resources._16x16);
            Image image32x32 = ByteArrayToImage(Resources._32x32);
            Image checkIn16x16 = ByteArrayToImage(Resources.CheckIn_16x16);
            Image checkIn32x32 = ByteArrayToImage(Resources.CheckIn_32x32);
            Image checkOut16x16 = ByteArrayToImage(Resources.CheckOut_16x16);
            Image checkOut32x32 = ByteArrayToImage(Resources.CheckOut_32x32);
            Image undoCheckout16x16 = ByteArrayToImage(Resources.UndoCheckout_16x16);
            Image undoCheckout32x32 = ByteArrayToImage(Resources.UndoCheckout_32x32);
            Image copy16x16 = ByteArrayToImage(Resources.Copy_16x16);
            Image copy32x32 = ByteArrayToImage(Resources.Copy_32x32);
            Image get16x16 = ByteArrayToImage(Resources.Get_16x16);
            Image get32x32 = ByteArrayToImage(Resources.Get_32x32);
            Image overwrite16x16 = ByteArrayToImage(Resources.Overwrite_16x16);
            Image overwrite32x32 = ByteArrayToImage(Resources.Overwrite_32x32);
            Image refresh16x16 = ByteArrayToImage(Resources.Refresh_16x16);
            Image refresh32x32 = ByteArrayToImage(Resources.Refresh_32x32);





            if (!Utilities.ButtonDefExist("CustomRefresh"))
            {
                Refresh = Utilities.CreateButtonDef("Refresh", "CustomRefresh", "", refresh16x16, refresh32x32);

                if (Utilities.ButtonDefExist("CustomRefresh"))
                {
                    manageRibbonPanel.CommandControls.AddButton(Refresh, true);
                    Refresh.OnExecute += Refresh_OnExecute;
                }
            }

            if (!Utilities.ButtonDefExist("GetLatestAllRules"))
            {
                GetLatestAllRules = Utilities.CreateButtonDef("Get Latest", "GetLatestAllRules", "", get16x16, get32x32);

                if (Utilities.ButtonDefExist("GetLatestAllRules"))
                {
                    allRulesRibbonPanel.CommandControls.AddButton(GetLatestAllRules);
                    GetLatestAllRules.OnExecute += GetLatestAllRules_OnExecute;
                }
                    
            }

            if (!Utilities.ButtonDefExist("GetLatestReleasedVersionAllRules"))
            {
                GetLatestReleasedAllRules = Utilities.CreateButtonDef("Get Latest Released Version", "GetLatestReleasedVersionAllRules", "", get16x16, get32x32);

                if (Utilities.ButtonDefExist("GetLatestReleasedVersionAllRules"))
                {
                    allRulesRibbonPanel.CommandControls.AddButton(GetLatestReleasedAllRules);
                    GetLatestReleasedAllRules.OnExecute += GetLatestReleasedAllRules_OnExecute;
                }
                 
            }

            if (!Utilities.ButtonDefExist("VaultGet"))
            {
                Get = Utilities.CreateButtonDef("Get", "VaultGet", "", get16x16, get32x32);

                if (Utilities.ButtonDefExist("VaultGet"))

                {
                    statusRibbonPanel.CommandControls.AddButton(Get, true);
                    Get.OnExecute += Get_OnExecute;
                }
                   
            }

            if (!Utilities.ButtonDefExist("Checkin"))
            {
                CheckIn = Utilities.CreateButtonDef("Check In", "Checkin", "", checkIn16x16, checkIn32x32);
                if (Utilities.ButtonDefExist("Checkin"))
                {
                    statusRibbonPanel.CommandControls.AddButton(CheckIn, true);
                    CheckIn.OnExecute += CheckIn_OnExecute;
                }
                    
            }

            if(!Utilities.ButtonDefExist("Checkout"))
            {
                CheckOut = Utilities.CreateButtonDef("Check Out", "Checkout", "", checkOut16x16, checkOut32x32);

                if (Utilities.ButtonDefExist("Checkout"))
                {
                    statusRibbonPanel.CommandControls.AddButton(CheckOut, true);
                    CheckOut.OnExecute += CheckOut_OnExecute;
                }
                  
            }

            if (!Utilities.ButtonDefExist("UndoCheckOut"))
            {
                UndoCheckOut = Utilities.CreateButtonDef("Undo Check Out", "UndoCheckOut", "", undoCheckout16x16, undoCheckout32x32);

                if (Utilities.ButtonDefExist("UndoCheckOut"))
                {
                    statusRibbonPanel.CommandControls.AddButton(UndoCheckOut, true);
                    UndoCheckOut.OnExecute += UndoCheckOut_OnExecute;
                }
                   
            }

            if (!Utilities.ComboExist("CurrentLifecycleState"))
            {
                CurrentLifecycleState = Utilities.CreateComboBoxDef("Current Lifecycle State", "CurrentLifecycleState", CommandTypesEnum.kQueryOnlyCmdType, 200);

                if (Utilities.ComboExist("CurrentLifecycleState"))
                {
                    CurrentLifecycleState.Enabled = false;
                    lifecycleRibbonPanel.CommandControls.AddComboBox(CurrentLifecycleState);
                    string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
                    if (ExternalRules.ListIndex != 1)
                    {
                        ResetCurrentLifecycleDropdown();
                    }
                }
            }

            if (!Utilities.ButtonDefExist("MakeLocalCopy"))
            {
                MakeLocalCopy = Utilities.CreateButtonDef("Make Local Copy", "MakeLocalCopy", "", copy16x16, copy32x32);

                if (Utilities.ButtonDefExist("MakeLocalCopy"))
                {
                    LocalRibbonPanel.CommandControls.AddButton(MakeLocalCopy, true);
                    MakeLocalCopy.OnExecute += MakeLocalCopy_OnExecute;
                }
                   
            }
      
            if (!Utilities.ButtonDefExist("OverwriteRuleOnDiskWithCopy"))
            {
                OverwriteRuleOnDiskWithCopy = Utilities.CreateButtonDef("Overwrite Rule\r\nWith Copy", "OverwriteRuleOnDiskWithCopy", "", overwrite16x16, overwrite32x32);

                if (Utilities.ButtonDefExist("OverwriteRuleOnDiskWithCopy"))
                {
                    LocalRibbonPanel.CommandControls.AddButton(OverwriteRuleOnDiskWithCopy, true);
                    OverwriteRuleOnDiskWithCopy.OnExecute += OverwriteRuleOnDiskWithCopy_OnExecute;
                }
                    
            }
        }


        public void ResetCurrentLifecycleDropdown()
        {
            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            CurrentLifecycleState.Clear();
            CurrentLifecycleState.AddItem(VaultLifecycleUtilities.GetLifecycleState(selectedItemName));
            CurrentLifecycleState.ListIndex = 1;
        }

        #endregion


    }
}