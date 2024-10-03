using ExternalRuleManager.Properties;
using Inventor;
using System.Data;
using ACW = Autodesk.Connectivity.WebServices;
using Autodesk.iLogic.Interfaces;

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
        public static ComboBoxDefinition LocalRules { get; set; }


        public ButtonDefinition GetLatestAllRules { get; set; }
        public ButtonDefinition GetLatestReleasedAllRules { get; set; }
        public ButtonDefinition Get { get; set; }
        public ButtonDefinition CheckIn { get; set; }
        public ButtonDefinition CheckOut { get; set; }
        public ButtonDefinition UndoCheckOut { get; set; }
        public ButtonDefinition MakeLocalCopy { get; set; }
        public ButtonDefinition OverwriteRuleOnDiskWithCopy { get; set; }
        public static ButtonDefinition Refresh { get; set; }
        public static ButtonDefinition RunRule { get; set; }


        public bool isExternalRulesInitialized = false;

        private RibbonTab externalRuleManagerRibbonZeroDocTab;
        private RibbonTab externalRuleManagerRibbonPartTab;
        private RibbonTab externalRuleManagerRibbonAssyTab;
        private RibbonTab externalRuleManagerRibbonDwgTab;
        


        private RibbonPanel manageRibbonPanel;
        private RibbonPanel statusRibbonPanel;
        private RibbonPanel allRulesRibbonPanel;
        private RibbonPanel lifecycleRibbonPanel;
        private RibbonPanel LocalRibbonPanel;
        private RibbonPanel localRulesRibbon;

        private System.Drawing.Image image16x16;
        private System.Drawing.Image image32x32;
        private System.Drawing.Image checkIn16x16;
        private System.Drawing.Image checkIn32x32;
        private System.Drawing.Image checkOut16x16;
        private System.Drawing.Image checkOut32x32;
        private System.Drawing.Image undoCheckout16x16;
        private System.Drawing.Image undoCheckout32x32;
        private System.Drawing.Image copy16x16;
        private System.Drawing.Image copy32x32;
        private System.Drawing.Image get16x16;
        private System.Drawing.Image get32x32;
        private System.Drawing.Image overwrite16x16;
        private System.Drawing.Image overwrite32x32;
        private System.Drawing.Image refresh16x16;
        private System.Drawing.Image refresh32x32;
        private System.Drawing.Image runRule16x16;
        private System.Drawing.Image runRule32x32;


        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomRibbon"/> class.
        /// Creates various buttons and ComboBoxes and adds them to the specified ribbons in the Inventor UI.
        /// </summary>
        public CustomRibbon()
        {

            // Load images for the ribbon buttons from resources
            image16x16 = ByteArrayToImage(Resources._16x16);
            image32x32 = ByteArrayToImage(Resources._32x32);
            checkIn16x16 = ByteArrayToImage(Resources.CheckIn_16x16);
            checkIn32x32 = ByteArrayToImage(Resources.CheckIn_32x32);
            checkOut16x16 = ByteArrayToImage(Resources.CheckOut_16x16);
            checkOut32x32 = ByteArrayToImage(Resources.CheckOut_32x32);
            undoCheckout16x16 = ByteArrayToImage(Resources.UndoCheckout_16x16);
            undoCheckout32x32 = ByteArrayToImage(Resources.UndoCheckout_32x32);
            copy16x16 = ByteArrayToImage(Resources.Copy_16x16);
            copy32x32 = ByteArrayToImage(Resources.Copy_32x32);
            get16x16 = ByteArrayToImage(Resources.Get_16x16);
            get32x32 = ByteArrayToImage(Resources.Get_32x32);
            overwrite16x16 = ByteArrayToImage(Resources.Overwrite_16x16);
            overwrite32x32 = ByteArrayToImage(Resources.Overwrite_32x32);
            refresh16x16 = ByteArrayToImage(Resources.Refresh_16x16);
            refresh32x32 = ByteArrayToImage(Resources.Refresh_32x32);
            runRule16x16 = ByteArrayToImage(Resources.RunRule_16x16);
            runRule32x32 = ByteArrayToImage(Resources.RunRule_32x32);

            // Create ComboBoxes and Buttons
            ExternalRules = Utilities.CreateComboBoxDef("External Rules", "ExternalRules", CommandTypesEnum.kQueryOnlyCmdType, 200);
            ExternalRules.OnSelect += ExternalRules_OnSelect;

            Refresh = Utilities.CreateButtonDef("Refresh", "CustomRefresh", "", refresh16x16, refresh32x32);
            Refresh.OnExecute += Refresh_OnExecute;

            GetLatestAllRules = Utilities.CreateButtonDef("Get Latest", "GetLatestAllRules", "", get16x16, get32x32);
            GetLatestAllRules.OnExecute += GetLatestAllRules_OnExecute;

            GetLatestReleasedAllRules = Utilities.CreateButtonDef("Get Latest Released Version", "GetLatestReleasedVersionAllRules", "", get16x16, get32x32);
            GetLatestReleasedAllRules.OnExecute += GetLatestReleasedAllRules_OnExecute;

            Get = Utilities.CreateButtonDef("Get", "VaultGet", "", get16x16, get32x32);
            Get.OnExecute += Get_OnExecute;

            CheckIn = Utilities.CreateButtonDef("Check In", "Checkin", "", checkIn16x16, checkIn32x32);
            CheckIn.OnExecute += CheckIn_OnExecute;

            CheckOut = Utilities.CreateButtonDef("Check Out", "Checkout", "", checkOut16x16, checkOut32x32);
            CheckOut.OnExecute += CheckOut_OnExecute;

            UndoCheckOut = Utilities.CreateButtonDef("Undo Check Out", "UndoCheckOut", "", undoCheckout16x16, undoCheckout32x32);
            UndoCheckOut.OnExecute += UndoCheckOut_OnExecute;

            CurrentLifecycleState = Utilities.CreateComboBoxDef("Current Lifecycle State", "CurrentLifecycleState", CommandTypesEnum.kQueryOnlyCmdType, 200);
            CurrentLifecycleState.Enabled = false;

            MakeLocalCopy = Utilities.CreateButtonDef("Make Local Copy", "MakeLocalCopy", "", copy16x16, copy32x32);
            MakeLocalCopy.OnExecute += MakeLocalCopy_OnExecute;

            OverwriteRuleOnDiskWithCopy = Utilities.CreateButtonDef("Overwrite Rule\r\nWith Copy", "OverwriteRuleOnDiskWithCopy", "", overwrite16x16, overwrite32x32);
            OverwriteRuleOnDiskWithCopy.OnExecute += OverwriteRuleOnDiskWithCopy_OnExecute;

            RunRule = Utilities.CreateButtonDef("Run Rule", "RunRule", "", runRule16x16, runRule32x32);
            RunRule.OnExecute += RunRule_OnExecute;

            // Add to specific ribbons
            Ribbon partRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            Ribbon assemblyRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            Ribbon drawingRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            Ribbon zeroDocRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["ZeroDoc"];

            AddToRibbon(partRibbon);
            AddToRibbon(assemblyRibbon);
            AddToRibbon(drawingRibbon);
            AddToRibbon(zeroDocRibbon);

            UIDisable(Globals.InvApp.ActiveDocument);
        }

        /// <summary>
        /// Converts a byte array to an Image object.
        /// </summary>
        /// <param name="byteArray">The byte array to convert.</param>
        /// <returns>An Image object created from the byte array.</returns>
        private System.Drawing.Image ByteArrayToImage(byte[] byteArray)
        {
            using (MemoryStream ms = new MemoryStream(byteArray))
            {
                return System.Drawing.Image.FromStream(ms);
            }
        }

        /// <summary>
        /// Adds the CustomRibbon buttons and controls to the specified ribbon in the Inventor UI.
        /// </summary>
        /// <param name="targetRibbon">The ribbon to which the CustomRibbon controls will be added.</param>
        public void AddToRibbon(Ribbon targetRibbon)
        {
            // Create and add ribbon tabs and panels
            externalRuleManagerRibbonZeroDocTab = targetRibbon.RibbonTabs.Add("External Rule Manager", "id_ExternalRuleManagerZeroDocTab" + targetRibbon.InternalName, Globals.InvAppGuidID);

            manageRibbonPanel = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("Manage","id_Manage" + targetRibbon.InternalName, Globals.InvAppGuidID);

            allRulesRibbonPanel = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("All Rules","id_AllRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            statusRibbonPanel = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("Status", "id_Status" + targetRibbon.InternalName, Globals.InvAppGuidID);

            lifecycleRibbonPanel = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("Lifecycle","id_Lifecycle" + targetRibbon.InternalName, Globals.InvAppGuidID);

            LocalRibbonPanel = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("Local (C:)","id_Local" + targetRibbon.InternalName, Globals.InvAppGuidID);

            localRulesRibbon = externalRuleManagerRibbonZeroDocTab.RibbonPanels.Add("Local Rules","id_LocalRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            // Add command controls to ribbon panels
            manageRibbonPanel.CommandControls.AddComboBox(ExternalRules);
            manageRibbonPanel.CommandControls.AddButton(Refresh, true);
            manageRibbonPanel.CommandControls.AddButton(RunRule, true);
            allRulesRibbonPanel.CommandControls.AddButton(GetLatestAllRules);
            allRulesRibbonPanel.CommandControls.AddButton(GetLatestReleasedAllRules);
            statusRibbonPanel.CommandControls.AddButton(Get, true);
            statusRibbonPanel.CommandControls.AddButton(CheckIn, true);
            statusRibbonPanel.CommandControls.AddButton(CheckOut, true);
            statusRibbonPanel.CommandControls.AddButton(UndoCheckOut, true);
            lifecycleRibbonPanel.CommandControls.AddComboBox(CurrentLifecycleState);
            LocalRibbonPanel.CommandControls.AddButton(MakeLocalCopy, true);
            LocalRibbonPanel.CommandControls.AddButton(OverwriteRuleOnDiskWithCopy, true);
        }


        #endregion


        #region Ui Events

        /// <summary>
        /// Executes the Refresh command, updating the user interface.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void Refresh_OnExecute(NameValueMap context)
        {
            RefreshUI();
        }

        /// <summary>
        /// Handles the selection of external rules and refreshes the user interface.
        /// </summary>
        /// <param name="context">The context for the current selection, provided by NameValueMap.</param>
        private void ExternalRules_OnSelect(NameValueMap context)
        {
            RefreshUI();
        }

        /// <summary>
        /// Retrieves the latest version of all external rules and refreshes the user interface.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void GetLatestAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestOnFolder(Globals.ExternalRuleName);
            RefreshUI();
        }

        /// <summary>
        /// Retrieves the latest "Released" version of all external rules and refreshes the user interface.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void GetLatestReleasedAllRules_OnExecute(NameValueMap context)
        {
            VaultUtilities.GetLatestFilesByLifecycleState(Globals.ExternalRuleName, "Released");
            RefreshUI();
        }

        /// <summary>
        /// Undoes the checkout for the selected file in the external rules ComboBox and refreshes the UI.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void UndoCheckOut_OnExecute(NameValueMap context)
        {
            ComboBoxDefinition tempComboBox = ExternalRules;

            if (tempComboBox != null)
            {
                string selectedItemName = tempComboBox.ListItem[tempComboBox.ListIndex];
                if (tempComboBox.ListIndex != 1)
                {
                    VaultFileUtilities.File_UndoCheckOut(selectedItemName);
                }

                RefreshUI();
            }
        }

        /// <summary>
        /// Retrieves the latest version of the selected file from the external rules ComboBox and refreshes the UI.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void Get_OnExecute(NameValueMap context)
        {
            ComboBoxDefinition tempComboBox = ExternalRules;

            if (tempComboBox != null)
            {
                string selectedItemName = tempComboBox.ListItem[tempComboBox.ListIndex];
                if (tempComboBox.ListIndex != 1)
                {
                    ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);
                    VaultFileUtilities.File_GetLatest(selectedItemName);
                }

                RefreshUI();
            }
        }

        /// <summary>
        /// Checks out the selected file from the external rules ComboBox and refreshes the UI.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void CheckOut_OnExecute(NameValueMap context)
        {
            ComboBoxDefinition tempComboBox = ExternalRules;

            if (tempComboBox != null)
            {
                string selectedItemName = tempComboBox.ListItem[tempComboBox.ListIndex];
                if (tempComboBox.ListIndex != 1)
                {
                    ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);
                    VaultFileUtilities.File_CheckOut(file);
                }

                RefreshUI();
            }
        }

        /// <summary>
        /// Checks in the selected file from the external rules ComboBox and refreshes the UI.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void CheckIn_OnExecute(NameValueMap context)
        {
            ComboBoxDefinition tempComboBox = ExternalRules;

            if (tempComboBox != null)
            {
                string selectedItemName = tempComboBox.ListItem[tempComboBox.ListIndex];
                if (tempComboBox.ListIndex != 1)
                {
                    ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);
                    VaultFileUtilities.File_CheckIn(selectedItemName, "");
                }

                RefreshUI();
            }
        }

        /// <summary>
        /// Retrieves the lifecycle state of the selected file from the external rules ComboBox.
        /// </summary>
        /// <param name="context">The context for the current selection, provided by NameValueMap.</param>
        private void CurrentLifecycleState_OnSelect(NameValueMap context)
        {
            ComboBoxDefinition tempComboBox = ExternalRules;

            if (tempComboBox != null)
            {
                string selectedItemName = tempComboBox.ListItem[tempComboBox.ListIndex];
                if (tempComboBox.ListIndex != 1)
                {
                    VaultLifecycleUtilities.GetLifecycleState(selectedItemName);
                }
            }
        }

        /// <summary>
        /// Makes a local copy of the external rules file and refreshes the user interface.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void MakeLocalCopy_OnExecute(NameValueMap context)
        {
            FileUtilities.MakeLocalCopy();
            RefreshUI();
        }

        /// <summary>
        /// Overwrites the external rule on the disk with the local copy and refreshes the user interface.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void OverwriteRuleOnDiskWithCopy_OnExecute(NameValueMap context)
        {
            FileUtilities.OverwriteRuleOnDiskWithCopy();
            RefreshUI();
        }

        /// <summary>
        /// Executes the selected external rule for the active document.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void RunRule_OnExecute(NameValueMap context)
        {
            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            Document activeDoc = Globals.InvApp.ActiveDocument;

            IStandardObjectProvider SOP = StandardObjectFactory.Create(Globals.InvApp.ActiveDocument);

            SOP.iLogicVb.RunExternalRule(Globals.ExternalRuleDir + "\\" + selectedItemName);
        }

        /// <summary>
        /// Resets the ribbon interface and displays a message indicating the ribbon or tab has changed.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by NameValueMap.</param>
        private void UserInterfaceManager_OnResetRibbonInterface(Inventor.NameValueMap context)
        {
            // Code to execute when the ribbon interface is reset or a tab is changed
            MessageBox.Show("Ribbon tab or interface has changed!");
        }




        #endregion


        #region Methods

        /// <summary>
        /// Adds external rules to the ExternalRules ComboBox by retrieving files from Vault,
        /// sorting them alphabetically, and setting the first item as "No External Rule Selected".
        /// </summary>
        public void AddExternalRules()
        {
            VaultConn.InitializeConnection();
            ACW.File[]? files = VaultUtilities.GetFilesByFolder(Globals.ExternalRuleName);

            if (files != null && files.Length > 0)
            {
                if (ExternalRules.ListIndex != 1)
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
        }

        /// <summary>
        /// Refreshes the user interface based on the currently selected item in the ExternalRules ComboBox.
        /// Enables or disables buttons based on whether a rule is selected and checked out.
        /// </summary>
        public void RefreshUI()
        {
            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            bool isCheckedOut = false;
            string currentRibbon = Globals.InvApp.UserInterfaceManager.ActiveEnvironment.Ribbon.InternalName;

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
                RunRule.Enabled = false;
            }
            else
            {
                if (ExternalRules.ListIndex != 1)
                {
                    ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                    if (file != null)
                    {
                        isCheckedOut = VaultFileUtilities.File_IsCheckedOut(file.Name);
                        Refresh.Enabled = true;
                        GetLatestAllRules.Enabled = true;
                        GetLatestReleasedAllRules.Enabled = true;
                        Get.Enabled = true;
                        CheckOut.Enabled = !isCheckedOut;
                        UndoCheckOut.Enabled = isCheckedOut;
                        CheckIn.Enabled = isCheckedOut;
                        MakeLocalCopy.Enabled = true;
                        OverwriteRuleOnDiskWithCopy.Enabled = true;
                        RunRule.Enabled = currentRibbon != "ZeroDoc";
                        ResetCurrentLifecycleDropdown();
                    }
                }
            }
        }

        /// <summary>
        /// Resets the CurrentLifecycleState dropdown based on the selected item in the ExternalRules ComboBox.
        /// </summary>
        public void ResetCurrentLifecycleDropdown()
        {
            string selectedItemName = ExternalRules.ListItem[ExternalRules.ListIndex];
            CurrentLifecycleState.Clear();
            CurrentLifecycleState.AddItem(VaultLifecycleUtilities.GetLifecycleState(selectedItemName));
            CurrentLifecycleState.ListIndex = 1;
        }

        /// <summary>
        /// Resets the LocalRules dropdown by clearing the existing items and adding sorted file names
        /// from the local directory. Sets the first item as selected.
        /// </summary>
        /// <exception cref="ArgumentException">Thrown when the local directory is not found.</exception>
        public void ResetLocalRulesDropdown()
        {
            LocalRules.Clear();

            try
            {
                string[] files = Directory.GetFiles(Globals.ExternalRuleDir)
                                          .OrderBy(file => System.IO.Path.GetFileName(file))
                                          .ToArray();

                foreach (string file in files)
                {
                    LocalRules.AddItem(System.IO.Path.GetFileName(file));
                }

                LocalRules.ListIndex = 1;
            }
            catch (Exception)
            {
                throw new ArgumentException("Directory not found...");
            }
        }

        /// <summary>
        /// Loads both Vault and local rules, combines them, removes duplicates, and sorts them alphabetically.
        /// Populates the ExternalRules ComboBox with the combined list of rules.
        /// </summary>
        /// <exception cref="ArgumentException">Thrown when the local directory is not found.</exception>
        public void LoadRules()
        {
            VaultConn.InitializeConnection();
            ACW.File[]? vaultFiles = VaultUtilities.GetFilesByFolder(Globals.ExternalRuleName);

            string[] localFiles = Array.Empty<string>();
            try
            {
                localFiles = Directory.GetFiles(Globals.ExternalRuleDir);
            }
            catch (Exception)
            {
                throw new ArgumentException("Local directory not found...");
            }

            List<string> combinedFiles = new List<string>();

            if (vaultFiles != null && vaultFiles.Length > 0)
            {
                combinedFiles.AddRange(vaultFiles.Select(file => file.Name));
            }

            combinedFiles.AddRange(localFiles.Select(file => System.IO.Path.GetFileName(file)));

            List<string> distinctSortedFiles = combinedFiles.Distinct().OrderBy(fileName => fileName).ToList();

            ExternalRules.Clear();
            ExternalRules.AddItem("No External Rule Selected");

            foreach (string fileName in distinctSortedFiles)
            {
                ExternalRules.AddItem(fileName);
            }

            ExternalRules.ListIndex = 1;
        }

        /// <summary>
        /// Enables specific user interface elements for external rules based on the given document.
        /// </summary>
        /// <param name="document">The current active document in Inventor.</param>
        public void UIEnable(Inventor.Document document)
        {
            ExternalRules.Enabled = true;
            Refresh.Enabled = false;
            GetLatestAllRules.Enabled = false;
            GetLatestReleasedAllRules.Enabled = false;
            Get.Enabled = false;
            CheckIn.Enabled = false;
            CheckOut.Enabled = false;
            UndoCheckOut.Enabled = false;
            CurrentLifecycleState.Enabled = false;
            MakeLocalCopy.Enabled = false;
            OverwriteRuleOnDiskWithCopy.Enabled = false;
            RunRule.Enabled = false;
        }

        /// <summary>
        /// Disables all user interface elements for external rules.
        /// </summary>
        /// <param name="document">The current active document in Inventor.</param>
        public void UIDisable(Inventor.Document document)
        {
            ExternalRules.Enabled = false;
            Refresh.Enabled = false;
            GetLatestAllRules.Enabled = false;
            GetLatestReleasedAllRules.Enabled = false;
            Get.Enabled = false;
            CheckIn.Enabled = false;
            CheckOut.Enabled = false;
            UndoCheckOut.Enabled = false;
            CurrentLifecycleState.Enabled = false;
            MakeLocalCopy.Enabled = false;
            OverwriteRuleOnDiskWithCopy.Enabled = false;
            RunRule.Enabled = false;
        }







        #endregion


    }
}