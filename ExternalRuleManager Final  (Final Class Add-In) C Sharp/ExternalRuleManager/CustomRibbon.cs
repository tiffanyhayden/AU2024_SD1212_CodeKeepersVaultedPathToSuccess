using ExternalRuleManager.Properties;
using Inventor;
using System.Data;
using ACW = Autodesk.Connectivity.WebServices;
using Autodesk.iLogic.Interfaces;
using Autodesk.iLogic.Automation;

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
        public ComboBoxDefinition LocalRules { get; set; }


        public ButtonDefinition GetLatestAllRules { get; set; }
        public ButtonDefinition GetLatestReleasedAllRules { get; set; }
        public ButtonDefinition Get { get; set; }
        public ButtonDefinition CheckIn { get; set; }
        public ButtonDefinition CheckOut { get; set; }
        public ButtonDefinition UndoCheckOut { get; set; }
        public ButtonDefinition MakeLocalCopy { get; set; }
        public ButtonDefinition OverwriteRuleOnDiskWithCopy { get; set; }
        public ButtonDefinition Refresh { get; set; }
        public ButtonDefinition RunRule { get; set; }

        public ButtonDefinition SetExternalRulePath { get; set; }   

        private RibbonTab _externalRuleManagerRibbonTab;
        


        private RibbonPanel _manageRibbonPanel;
        private RibbonPanel _statusRibbonPanel;
        private RibbonPanel _allRulesRibbonPanel;
        private RibbonPanel _lifecycleRibbonPanel;
        private RibbonPanel _localRibbonPanel;
        private RibbonPanel _localRulesRibbon;

        private System.Drawing.Image _image16x16;
        private System.Drawing.Image _image32x32;
        private System.Drawing.Image _checkIn16x16;
        private System.Drawing.Image _checkIn32x32;
        private System.Drawing.Image _checkOut16x16;
        private System.Drawing.Image _checkOut32x32;
        private System.Drawing.Image _undoCheckout16x16;
        private System.Drawing.Image _undoCheckout32x32;
        private System.Drawing.Image _copy16x16;
        private System.Drawing.Image _copy32x32;
        private System.Drawing.Image _get16x16;
        private System.Drawing.Image _get32x32;
        private System.Drawing.Image _overwrite16x16;
        private System.Drawing.Image _overwrite32x32;
        private System.Drawing.Image _refresh16x16;
        private System.Drawing.Image _refresh32x32;
        private System.Drawing.Image _runRule16x16;
        private System.Drawing.Image _runRule32x32;
        private System.Drawing.Image _setRule16x16;
        private System.Drawing.Image _setRule32x32;

        private bool IsExternalRuleSelected()
        {
            return ExternalRules.ListIndex != 1;
        }


        private void LoadButtonImages()
        {
            // Load images for the ribbon buttons from resources
            _image16x16 = LoadImage(Resources._16x16);
            _image32x32 = LoadImage(Resources._32x32);
            _checkIn16x16 = LoadImage(Resources.CheckIn_16x16);
            _checkIn32x32 = LoadImage(Resources.CheckIn_32x32);
            _checkOut16x16 = LoadImage(Resources.CheckOut_16x16);
            _checkOut32x32 = LoadImage(Resources.CheckOut_32x32);
            _undoCheckout16x16 = LoadImage(Resources.UndoCheckout_16x16);
            _undoCheckout32x32 = LoadImage(Resources.UndoCheckout_32x32);
            _copy16x16 = LoadImage(Resources.Copy_16x16);
            _copy32x32 = LoadImage(Resources.Copy_32x32);
            _get16x16 = LoadImage(Resources.Get_16x16);
            _get32x32 = LoadImage(Resources.Get_32x32);
            _overwrite16x16 = LoadImage(Resources.Overwrite_16x16);
            _overwrite32x32 = LoadImage(Resources.Overwrite_32x32);
            _refresh16x16 = LoadImage(Resources.Refresh_16x16);
            _refresh32x32 = LoadImage(Resources.Refresh_32x32);
            _runRule16x16 = LoadImage(Resources.RunRule_16x16);
            _runRule32x32 = LoadImage(Resources.RunRule_32x32);
            _setRule16x16 = LoadImage(Resources.Settings_16x16);
            _setRule32x32 = LoadImage(Resources.Settings_32x32);
        }

        private System.Drawing.Image LoadImage(byte[] resource)
        {
            return ByteArrayToImage(resource);
        }


        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomRibbon"/> class.
        /// Creates various buttons and ComboBoxes and adds them to the specified ribbons in the Inventor UI.
        /// </summary>
        public CustomRibbon()
        {
            LoadButtonImages();


            // Create ComboBoxes and Buttons
            ExternalRules = Utilities.CreateComboBoxDef("External Rules", "ExternalRules", CommandTypesEnum.kQueryOnlyCmdType, 200);
            ExternalRules.OnSelect += ExternalRules_OnSelect;


            SetExternalRulePath = Utilities.CreateButtonDef("Set External Rule Path", "SetExternalRulePath", "", _setRule16x16, _setRule32x32);
            SetExternalRulePath.OnExecute += SetExternalRulePath_OnExecute;

            Refresh = Utilities.CreateButtonDef("Refresh", "CustomRefresh", "", _refresh16x16, _refresh32x32);
            Refresh.OnExecute += Refresh_OnExecute;

            GetLatestAllRules = Utilities.CreateButtonDef("Get Latest", "GetLatestAllRules", "", _get16x16, _get32x32);
            GetLatestAllRules.OnExecute += GetLatestAllRules_OnExecute;

            GetLatestReleasedAllRules = Utilities.CreateButtonDef("Get Latest Released Version", "GetLatestReleasedVersionAllRules", "", _get16x16, _get32x32);
            GetLatestReleasedAllRules.OnExecute += GetLatestReleasedAllRules_OnExecute;

            Get = Utilities.CreateButtonDef("Get", "VaultGet", "", _get16x16, _get32x32);
            Get.OnExecute += Get_OnExecute;

            CheckIn = Utilities.CreateButtonDef("Check In", "Checkin", "", _checkIn16x16, _checkIn32x32);
            CheckIn.OnExecute += CheckIn_OnExecute;

            CheckOut = Utilities.CreateButtonDef("Check Out", "Checkout", "", _checkOut16x16, _checkOut32x32);
            CheckOut.OnExecute += CheckOut_OnExecute;

            UndoCheckOut = Utilities.CreateButtonDef("Undo Check Out", "UndoCheckOut", "", _undoCheckout16x16, _undoCheckout32x32);
            UndoCheckOut.OnExecute += UndoCheckOut_OnExecute;

            CurrentLifecycleState = Utilities.CreateComboBoxDef("Current Lifecycle State", "CurrentLifecycleState", CommandTypesEnum.kQueryOnlyCmdType, 200);
            CurrentLifecycleState.Enabled = false;

            MakeLocalCopy = Utilities.CreateButtonDef("Make Local Copy", "MakeLocalCopy", "", _copy16x16, _copy32x32);
            MakeLocalCopy.OnExecute += MakeLocalCopy_OnExecute;

            OverwriteRuleOnDiskWithCopy = Utilities.CreateButtonDef("Overwrite Rule\r\nWith Copy", "OverwriteRuleOnDiskWithCopy", "", _overwrite16x16, _overwrite32x32);
            OverwriteRuleOnDiskWithCopy.OnExecute += OverwriteRuleOnDiskWithCopy_OnExecute;

            RunRule = Utilities.CreateButtonDef("Run Rule", "RunRule", "", _runRule16x16, _runRule32x32);
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

            Utilities.SetExternalRuleDirectory("");
        }

        /// <summary>
        /// Converts a byte array to an Image object.
        /// </summary>
        /// <param name="byteArray">The byte array to convert.</param>
        /// <returns>An Image object created from the byte array.</returns>
        private System.Drawing.Image ByteArrayToImage(byte[] byteArray)
        {
            using var ms = new MemoryStream(byteArray);
            return System.Drawing.Image.FromStream(ms);
        }

        /// <summary>
        /// Adds the CustomRibbon buttons and controls to the specified ribbon in the Inventor UI.
        /// </summary>
        /// <param name="targetRibbon">The ribbon to which the CustomRibbon controls will be added.</param>
        public void AddToRibbon(Ribbon targetRibbon)
        {
            // Create and add ribbon tabs and panels
            _externalRuleManagerRibbonTab = targetRibbon.RibbonTabs.Add("External Rule Manager", "id_ExternalRuleManagerZeroDocTab" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _manageRibbonPanel = _externalRuleManagerRibbonTab.RibbonPanels.Add("Manage","id_Manage" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _allRulesRibbonPanel = _externalRuleManagerRibbonTab.RibbonPanels.Add("All Rules","id_AllRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _statusRibbonPanel = _externalRuleManagerRibbonTab.RibbonPanels.Add("Status", "id_Status" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _lifecycleRibbonPanel = _externalRuleManagerRibbonTab.RibbonPanels.Add("Lifecycle","id_Lifecycle" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _localRibbonPanel = _externalRuleManagerRibbonTab.RibbonPanels.Add("Local (C:)","id_Local" + targetRibbon.InternalName, Globals.InvAppGuidID);

            _localRulesRibbon = _externalRuleManagerRibbonTab.RibbonPanels.Add("Local Rules","id_LocalRules" + targetRibbon.InternalName, Globals.InvAppGuidID);

            // Add command controls to ribbon panels
            _manageRibbonPanel.CommandControls.AddComboBox(ExternalRules);
            _manageRibbonPanel.CommandControls.AddButton(SetExternalRulePath);
            _manageRibbonPanel.CommandControls.AddButton(Refresh, true);
            _manageRibbonPanel.CommandControls.AddButton(RunRule, true);
            
            _allRulesRibbonPanel.CommandControls.AddButton(GetLatestAllRules);
            _allRulesRibbonPanel.CommandControls.AddButton(GetLatestReleasedAllRules);
            _statusRibbonPanel.CommandControls.AddButton(Get, true);
            _statusRibbonPanel.CommandControls.AddButton(CheckIn, true);
            _statusRibbonPanel.CommandControls.AddButton(CheckOut, true);
            _statusRibbonPanel.CommandControls.AddButton(UndoCheckOut, true);
            _lifecycleRibbonPanel.CommandControls.AddComboBox(CurrentLifecycleState);
            _localRibbonPanel.CommandControls.AddButton(MakeLocalCopy, true);
            _localRibbonPanel.CommandControls.AddButton(OverwriteRuleOnDiskWithCopy, true);

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

            LoadExternalRules();
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
        /// Executes the process of setting the path for external rules by prompting the user with a folder browser dialog.
        /// The selected folder path is then applied to the external rule directories.
        /// </summary>
        /// <param name="context">The context for the current execution, provided by <see cref="NameValueMap"/>.</param>
        /// <remarks>
        /// This method opens a <see cref="FolderBrowserDialog"/> to allow the user to choose a directory. 
        /// The chosen path is assigned to the application's external rule directories using the iLogic automation interface.
        /// Make sure that <c>Globals.InvApp.ActiveDocument</c> is properly initialized before invoking this method to avoid null reference exceptions.
        /// </remarks>
        private void SetExternalRulePath_OnExecute(NameValueMap context)
        {
            RefreshUI();

            // Create a new instance of the FolderBrowserDialog
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())


            {
                // Set the description that will be displayed in the dialog
                folderBrowserDialog.Description = "Select a folder";

                // Optionally set the initial directory if needed
                folderBrowserDialog.SelectedPath = "C:\\";

                // Show the dialog and check if the user clicked the OK button
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected folder path
                    string selectedPath = folderBrowserDialog.SelectedPath;

                    IStandardObjectProvider SOP = StandardObjectFactory.Create(Globals.InvApp.ActiveDocument);

                    SOP.iLogicVb.Automation.FileOptions.ExternalRuleDirectories = [selectedPath];

                    // Perform actions with the selected folder path
                    //MessageBox.Show($"Selected folder: {selectedPath}", "Folder Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

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
                if (IsExternalRuleSelected())
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
                if (IsExternalRuleSelected())
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
        public void LoadExternalRules()
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
            SetExternalRulePath.Enabled = true;
            Refresh.Enabled = true;
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


        private ComboBoxDefinition CreateNewComboBox(string displayName, string internalName, int width = 200)
        {
            var comboBox = Utilities.CreateComboBoxDef(displayName, internalName, CommandTypesEnum.kQueryOnlyCmdType, width);
            comboBox.OnSelect += (context) => RefreshUI();
            return comboBox;
        }

        private ButtonDefinition CreateNewButton(string displayName, string internalName, System.Drawing.Image icon16, System.Drawing.Image icon32)
        {
            var button = Utilities.CreateButtonDef(displayName, internalName, "", icon16, icon32);
            button.OnExecute += (context) => RefreshUI();  // Replace with actual handler as needed.
            return button;
        }




        #endregion


    }
}