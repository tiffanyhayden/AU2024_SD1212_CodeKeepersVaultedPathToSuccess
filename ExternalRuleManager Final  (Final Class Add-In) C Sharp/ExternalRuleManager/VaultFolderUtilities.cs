using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;

namespace ExternalRuleManager
{
    /// <summary>
    /// Provides utility methods for interacting with folders in Autodesk Vault.
    /// </summary>
    internal class VaultFolderUtilities
    {
        /// <summary>
        /// Acquires a Vault folder and downloads its contents to the local working folder.
        /// </summary>
        /// <param name="folder">The Vault folder to acquire.</param>
        /// <returns>The acquired folder, or <c>null</c> if the operation fails.</returns>
        /// <remarks>
        /// This method uses Vault connection settings to download folder contents, with options to manage file overwrites
        /// and relationship gathering settings.
        /// </remarks>
        public static ACW.Folder Folder_Acquire(ACW.Folder folder)
        {
            if (VaultConn.ActiveConnection != null)
            {
                IntPtr parent = IntPtr.Zero;
                VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings settings;
                VDF.Vault.Currency.Entities.IEntity? folderEntity = null;
                string workingFolder = VaultConn.ActiveConnection.WorkingFoldersManager.GetWorkingFolder("$").FullPath;

                string folderPath = folder.FullName;
                string updatedfolderPath = folderPath.Replace("/", @"\");
                string folderPathWithWorking = updatedfolderPath.Replace("$\\", workingFolder);

                // Ensure Vault is connected before proceeding.
                if (VaultConn.IsConnected())
                {
                    try
                    {
                        // Create the folder entity.
                        folderEntity = new VDF.Vault.Currency.Entities.Folder(VaultConn.ActiveConnection, folder);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    // Define settings for acquiring the folder.
                    settings = new VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings(VaultConn.ActiveConnection, parent, "Download files");
                    settings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                    settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                    settings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

                    // Set relationship gathering options.
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeHiddenEntities = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseParents = false;
                    settings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = false;

                    if (folderEntity != null)
                    {
                        try
                        {
                            settings.AddEntityToAcquire(folderEntity);
                            settings.LocalPath = new VDF.Currency.FolderPathAbsolute(folderPathWithWorking);
                            VaultConn.ActiveConnection?.FileManager.AcquireFiles(settings);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Acquiring folder failed.", "Error", MessageBoxButtons.OK);
                            throw;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Retrieves Vault folders based on their name.
        /// </summary>
        /// <param name="folderName">The name of the folder to search for.</param>
        /// <returns>An array of <see cref="ACW.Folder"/> objects matching the specified name, or <c>null</c> if none are found.</returns>
        public static ACW.Folder[]? GetFoldersByFolderName(string folderName)
        {
            ACW.Folder[]? folders = null;
            ACW.SrchCond searchCondition = new ACW.SrchCond();
            string bookmark = "";
            ACW.SrchStatus status;

            if (VaultConn.IsConnected())
            {
                Autodesk.Connectivity.WebServicesTools.WebServiceManager webServiceManager = VaultConn.ActiveConnection.WebServiceManager;
                ACW.DocumentService docService = webServiceManager.DocumentService;
                ACW.PropDef[] propDefs = webServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR");
                ACW.PropDef? propDef = Array.Find(propDefs, n => n.DispName == "Name");

                searchCondition.PropDefId = propDef.Id;
                searchCondition.PropTyp = ACW.PropertySearchType.SingleProperty;
                searchCondition.SrchOper = 3; // Search operator
                searchCondition.SrchRule = ACW.SearchRuleType.Must;
                searchCondition.SrchTxt = folderName;

                try
                {
                    folders = docService.FindFoldersBySearchConditions(new ACW.SrchCond[] { searchCondition }, null, new long[] { }, true, ref bookmark, out status);
                }
                catch (Exception)
                {
                    // No folders found
                    return null;
                }

                return folders;
            }

            return null;
        }
    }
}
