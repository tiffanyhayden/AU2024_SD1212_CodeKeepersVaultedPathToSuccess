using Autodesk.iLogic.Interfaces;
using Inventor;
using static ExternalRuleManager.ImageConverter;

namespace ExternalRuleManager
{
    /// <summary>
    /// Provides utility methods for creating and retrieving Inventor button and combo box definitions.
    /// </summary>
    public class Utilities
    {
        /// <summary>
        /// Creates a new button definition in Inventor, or returns an existing button definition if one already exists.
        /// </summary>
        /// <param name="DisplayName">The display name for the button.</param>
        /// <param name="InternalName">The internal name for the button. This must be unique.</param>
        /// <param name="ToolTip">The tooltip to be shown when hovering over the button (optional).</param>
        /// <param name="SmallIcon">An optional small icon (<see cref="Image"/>) for the button.</param>
        /// <param name="LargeIcon">An optional large icon (<see cref="Image"/>) for the button.</param>
        /// <returns>
        /// The <see cref="ButtonDefinition"/> object if successful, or the existing definition if one exists with the same internal name. Returns <c>null</c> if the operation fails.
        /// </returns>
        public static Inventor.ButtonDefinition CreateButtonDef(string DisplayName,
            string InternalName,
            string ToolTip = "",
            Image SmallIcon = null,
            Image LargeIcon = null)
        {
            Inventor.ButtonDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ButtonDefinition;
            }
            catch (Exception)
            {
                // Optionally log exception
            }

            if (existingDef != null)
            {
                // Optionally show a message to the user
                return existingDef;
            }

            IPictureDisp iPic16 = null;
            if (SmallIcon != null)
            {
                try
                {
                    iPic16 = ConvertImageToIPictureDisp(SmallIcon);
                }
                catch (Exception)
                {
                    // Optionally show a message to the user
                }
            }

            IPictureDisp iPic32 = null;
            if (LargeIcon != null)
            {
                try
                {
                    iPic32 = ConvertImageToIPictureDisp(LargeIcon);
                }
                catch (Exception)
                {
                    // Optionally show a message to the user
                }
            }

            try
            {
                Inventor.ControlDefinitions controlDefs = Globals.InvApp.CommandManager.ControlDefinitions;

                ButtonDefinition btnDef = controlDefs.AddButtonDefinition(DisplayName,
                                                                           InternalName,
                                                                           CommandTypesEnum.kShapeEditCmdType,
                                                                           Globals.InvAppGuidID,
                                                                           "",
                                                                           ToolTip,
                                                                           iPic16,
                                                                           iPic32);
                return btnDef;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Checks if a button definition with the specified internal name exists in Inventor.
        /// </summary>
        /// <param name="InternalName">The internal name of the button to check.</param>
        /// <returns><c>true</c> if the button definition exists; otherwise, <c>false</c>.</returns>
        public static bool ButtonDefExist(string InternalName)
        {
            Inventor.ButtonDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ButtonDefinition;
            }
            catch (Exception)
            {
                // Optionally log exception
            }

            return existingDef != null;
        }

        /// <summary>
        /// Retrieves an existing button definition based on the internal name.
        /// </summary>
        /// <param name="InternalName">The internal name of the button.</param>
        /// <returns>The <see cref="ButtonDefinition"/> object if it exists, or <c>null</c> if not found.</returns>
        public static ButtonDefinition GetButtonDef(string InternalName)
        {
            Inventor.ButtonDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ButtonDefinition;
            }
            catch (Exception)
            {
                // Optionally log exception
            }

            return existingDef;
        }

        /// <summary>
        /// Creates a new combo box definition in Inventor.
        /// </summary>
        /// <param name="DisplayName">The display name of the combo box.</param>
        /// <param name="InternalName">The internal name for the combo box. This must be unique.</param>
        /// <param name="Classification">The command classification for the combo box.</param>
        /// <param name="DropDownWidth">The width of the drop-down list in pixels.</param>
        /// <returns>
        /// The <see cref="ComboBoxDefinition"/> object if successful, or <c>null</c> if an existing combo box definition with the same internal name exists or the operation fails.
        /// </returns>
        public static Inventor.ComboBoxDefinition CreateComboBoxDef(string DisplayName,
            string InternalName,
            CommandTypesEnum Classification,
            int DropDownWidth)
        {
            Inventor.ComboBoxDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ComboBoxDefinition;
            }
            catch (Exception)
            {
                // Optionally log exception
            }

            if (existingDef != null)
            {
                // Optionally show a message to the user
                return null;
            }

            try
            {
                Inventor.ControlDefinitions controlDefs = Globals.InvApp.CommandManager.ControlDefinitions;

                ComboBoxDefinition cmbDef = controlDefs.AddComboBoxDefinition(DisplayName,
                                                                              InternalName,
                                                                              Classification,
                                                                              DropDownWidth, null, "", "", null, null, ButtonDisplayEnum.kAlwaysDisplayText);
                return cmbDef;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Checks if a combo box definition with the specified internal name exists in Inventor.
        /// </summary>
        /// <param name="InternalName">The internal name of the combo box to check.</param>
        /// <returns><c>true</c> if the combo box definition exists; otherwise, <c>false</c>.</returns>
        public static bool ComboExist(string InternalName)
        {
            Inventor.ComboBoxDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ComboBoxDefinition;
            }
            catch (Exception)
            {
                // Optionally log exception
            }

            return existingDef != null;
        }

        /// <summary>
        /// Sets the external rule directory for the application based on the provided path or a default path if available.
        /// If the specified path does not exist, it will be created.
        /// </summary>
        /// <param name="selectedPath">The path selected by the user to set as the external rule directory.</param>
        /// <remarks>
        /// This method retrieves the current external rule directories from the active document and checks their validity.
        /// - If the current directories are invalid or empty, it attempts to use a default path defined in <c>Globals.ExternalRuleDir</c>.
        /// - If the default path does not exist, the method will create it.
        /// - If the provided <paramref name="selectedPath"/> is not valid, it will also attempt to create that directory.
        /// - If neither path can be created, the user is prompted with a warning message.
        /// </remarks>
        public static void SetExternalRuleDirectory(string selectedPath)
        {
            // Get the Standard Object Provider (SOP) for the active document
            IStandardObjectProvider SOP = StandardObjectFactory.Create(Globals.InvApp.ActiveDocument);

            // Get the current external rule directories setting
            string[] currentExternalRuleDirs = SOP.iLogicVb.Automation.FileOptions.ExternalRuleDirectories;

            // Check if the array is empty or contains invalid paths
            if (currentExternalRuleDirs == null || !currentExternalRuleDirs.Any(Directory.Exists))
            {
                // If the array is empty or no valid paths are found, use the default path if it exists
                if (!Directory.Exists(Globals.ExternalRuleDir))
                {
                    // If the default path does not exist, create it
                    try
                    {
                        Directory.CreateDirectory(Globals.ExternalRuleDir);
                        SOP.iLogicVb.Automation.FileOptions.ExternalRuleDirectories = new string[] { Globals.ExternalRuleDir };
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to create the default directory: {ex.Message}",
                                        "Directory Creation Error",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    SOP.iLogicVb.Automation.FileOptions.ExternalRuleDirectories = new string[] { Globals.ExternalRuleDir };
                }
            }
            else if (!Directory.Exists(selectedPath))
            {
                // If the selected path is not valid, try creating it
                try
                {
                    Directory.CreateDirectory(selectedPath);
                    Globals.ExternalRuleDir = selectedPath;
                    SOP.iLogicVb.Automation.FileOptions.ExternalRuleDirectories = new string[] { Globals.ExternalRuleDir };
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to create the selected directory: {ex.Message}",
                                    "Directory Creation Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                // If the directory is already set, ensure it's valid
                bool isValidPath = currentExternalRuleDirs.Any(Directory.Exists);
                if (!isValidPath)
                {
                    MessageBox.Show("The currently set external rules directory is invalid. Please set a valid directory.",
                                    "Invalid Path",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
            }
        }
    }
}
