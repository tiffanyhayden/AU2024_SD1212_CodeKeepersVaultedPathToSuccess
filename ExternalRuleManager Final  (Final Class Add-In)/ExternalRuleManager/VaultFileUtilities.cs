using Autodesk.Common.IO;
using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Forms.Settings;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;
using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using System;
using System.IO;
using System.Xml.Linq;

namespace ExternalRuleManager
{
    internal class VaultFileUtilities
    {

        public static void File_Acquire(ACW.File file, bool doCheckOut = false)
        {
            IntPtr oParent = IntPtr.Zero;
            InteractiveAcquireFileSettings oSettings;
            FileIteration oFileIteration;


            if (VaultConn.ActiveConnection != null)
            {
                oFileIteration = new FileIteration(VaultConn.ActiveConnection, file);
                oSettings = new InteractiveAcquireFileSettings(VaultConn.ActiveConnection, oParent, "Download files");

                oSettings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                oSettings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeHiddenEntities = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseParents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = false;

                if (doCheckOut)
                {
                    oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }

                oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;

                oSettings.AddFileToAcquire(oFileIteration, oSettings.DefaultAcquisitionOption);
                VaultConn.ActiveConnection.FileManager.AcquireFiles(oSettings);
            }
        }


        public static void File_CheckOut(ACW.File file)
        {
            IntPtr parent = IntPtr.Zero;
            InteractiveAcquireFileSettings oSettings;
            FileIteration oFileIteration;


            if (VaultConn.ActiveConnection != null)
            {
                oFileIteration = new FileIteration(VaultConn.ActiveConnection, file);
                oSettings = new InteractiveAcquireFileSettings(VaultConn.ActiveConnection, parent, "Download files");

                oSettings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                oSettings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeHiddenEntities = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseParents = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = false;

                oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;

                oSettings.AddFileToAcquire(oFileIteration, oSettings.DefaultAcquisitionOption);
                VaultConn.ActiveConnection.FileManager.AcquireFiles(oSettings);
            }
        }

        public static string File_GetLatestLifecycleStateVersion(string fileName, string lifecycleStateName)
        {
            // THIS FUNCTION "GETS" THE LATEST VERSION OF A FILE THAT IS AT A SPECIFIED LIFECYCLE STATE
            try
            {
                lifecycleStateName = lifecycleStateName.ToUpper();

                // CHECK INPUTS
                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(lifecycleStateName))
                {
                    throw new ArgumentException("File name or lifecycle state cannot be null or empty.");
                }

                // Check Vault connection
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                // Get Vault services
                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;

                // Find file by name
                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new InvalidOperationException($"File '{fileName}' not found.");
                }

                // Get file versions
                ACW.File[] fileVersions = docService.GetFilesByMasterId(file.MasterId);
                if (fileVersions == null || fileVersions.Length == 0)
                {
                    throw new InvalidOperationException($"No versions found for file '{fileName}'.");
                }

                // Reverse the array to process latest version first
                Array.Reverse(fileVersions);

                // FIND THE LATEST VERSION OF A FILE AT THE SPECIFIED LIFECYCLE STATE
                ACW.File fileByLifecycleState = null;
                foreach (ACW.File fileVersion in fileVersions)
                {
                    if (fileVersion.FileLfCyc.LfCycStateName.ToUpper() == lifecycleStateName)
                    {
                        fileByLifecycleState = fileVersion;
                        break;
                    }
                }

                if (fileByLifecycleState == null)
                {
                    throw new InvalidOperationException($"No file found with lifecycle state '{lifecycleStateName}'.");
                }

                // Get file iteration and full file path
                FileIteration fileIteration = new FileIteration(VaultConn.ActiveConnection, fileByLifecycleState);
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                // Acquire the file from Vault
                File_Acquire(fileByLifecycleState, false);

                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                // Handle invalid arguments
                Console.WriteLine($"File_GetLatestLifecycleStateVersion Failed: {ex.Message}");
                return "";
            }
        }

        public static ACW.File File_FindByFileName(string fileName, int srchOperator = 3, string searchInStartingFolder = "$/")
        {
            try
            {
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                ACW.PropDef[] propDefs = VaultConn.ActiveConnection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");

                // Ensure property definitions were found
                if (propDefs == null || !propDefs.Any())
                {
                    throw new InvalidOperationException("No property definitions found for 'FILE'.");
                }

                // Find the specific property definition for "ClientFileName"
                ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "ClientFileName");
                if (propDef == null)
                {
                    throw new InvalidOperationException("Property definition for 'ClientFileName' not found.");
                }

                // Prepare the search condition
                ACW.SrchCond search = new ACW.SrchCond
                {
                    PropDefId = propDef.Id,
                    PropTyp = ACW.PropertySearchType.SingleProperty,
                    SrchOper = srchOperator,
                    SrchRule = ACW.SearchRuleType.Must,
                    SrchTxt = fileName
                };

                // Find folders by path
                ACW.Folder[] folders = docService.FindFoldersByPaths(new string[] { searchInStartingFolder });
                if (folders == null || folders.Length == 0)
                {
                    throw new InvalidOperationException($"No folders found at the specified path: {searchInStartingFolder}");
                }

                // Collect folder IDs
                long[] folderIDs = folders.Where(f => f.Id != -1).Select(f => f.Id).ToArray();
                if (folderIDs == null || folderIDs.Length == 0)
                {
                    throw new InvalidOperationException("No valid folder IDs found.");
                }

                // Execute the file search
                string bookmark = string.Empty;
                ACW.SrchStatus status;
                ACW.File[] results = docService.FindFilesBySearchConditions(new ACW.SrchCond[] { search }, null, folderIDs, true, true, ref bookmark, out status);

                // Ensure results are found
                if (results == null || results.Length == 0)
                {
                    throw new InvalidOperationException($"No files found matching the search criteria: {fileName}");
                }

                return results[0];  // Return the first result
            }
            catch (InvalidOperationException ex)
            {
                // Log or display specific Vault-related errors
                Console.WriteLine("File_FindByFileName Failed: " + ex.Message);
                return null;
            }
        }

        public static string File_UndoCheckOut(string fileName)
        {
            try
            {
                // Check if Vault connection is valid
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                // Check if the fileName is valid
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("File name must have a value to continue.", nameof(fileName));
                }

                // Find the file by name
                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new FileNotFoundException($"File '{fileName}' not found in Vault.");
                }

                // Get the file iteration and full file path
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                // Perform the undo checkout

                if(File_IsCheckedOut(file.Name))
                {
                   VaultConn.ActiveConnection.FileManager.UndoCheckoutFile(fileIteration);
                }
                

                // Return the full file path
                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                // Handle invalid file name
                Console.WriteLine($"File_UndoCheckOut Failed: {ex.Message}");
                return "";
            }
        }

        public static bool File_IsCheckedOut(string fileName, bool toCurrentUser = false)
        {
            try
            {
                // Check if Vault connection is valid
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                // Check if the fileName is valid
                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("File name must have a value to continue.", nameof(fileName));
                }

                // Find the file by name
                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new FileNotFoundException($"File '{fileName}' was not found in Vault.");
                }

                // Get the file iteration
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                // Check if the file is checked out
                if (!toCurrentUser)
                {
                    return fileIteration.IsCheckedOut;
                }
                else
                {
                    return fileIteration.IsCheckedOutToCurrentUser;
                }
            }
            catch (ArgumentException ex)
            {
                // Handle invalid file name input
                Console.WriteLine($"File_IsCheckedOut Failed: {ex.Message}");
                return false;
            }
        }

        public static  FileAssocParam[] File_GetAssociations(string fileName)
        {
            if (VaultConn.ActiveConnection != null)
            {

                ACW.File file = VaultFileUtilities.File_FindByFileName(fileName);
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                VDF.Vault.Settings.FileRelationshipGatheringSettings relationshipSettings = new VDF.Vault.Settings.FileRelationshipGatheringSettings
                {
                    IncludeAttachments = true,
                    IncludeChildren = true,
                    IncludeParents = true,
                    IncludeRelatedDocumentation = true
                };

                // Retrieve the collection of FileAssocLite
                IEnumerable<ACW.FileAssocLite> fileAssocLite = VaultConn.ActiveConnection.FileManager.GetFileAssociationLites(
                    new long[] { fileIteration.EntityIterationId },
                    relationshipSettings
                );

                // Create a list to store FileAssocParam objects
                List<FileAssocParam> fileAssocParams = new List<FileAssocParam>();

                // Check if the collection is not null
                if (fileAssocLite != null)
                {
                    // Iterate over the FileAssocLite collection
                    foreach (ACW.FileAssocLite thisFileAssocLite in fileAssocLite)
                    {
                        // Exclude the child file that matches the current file iteration
                        if (thisFileAssocLite.CldFileId != fileIteration.EntityIterationId)
                        {
                            // Create a new FileAssocParam object
                            FileAssocParam par = new FileAssocParam
                            {
                                CldFileId = thisFileAssocLite.CldFileId,
                                RefId = thisFileAssocLite.RefId,
                                Source = thisFileAssocLite.Source,
                                Typ = thisFileAssocLite.Typ,
                                ExpectedVaultPath = thisFileAssocLite.ExpectedVaultPath
                            };

                            // Add the FileAssocParam object to the list
                            fileAssocParams.Add(par);
                        }
                    }
                }

                // Convert the List<FileAssocParam> to an array and return it
                return fileAssocParams.ToArray();
            }

            return null;

        }

        public static string File_CheckIn(string fileName, string comment) 
        {

            if (VaultConn.ActiveConnection != null)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(fileName);
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                VDF.Currency.FilePathAbsolute? filePathAbs = GetFilePathByFile(fileName);

                if (!System.IO.File.Exists(filePathAbs.FileName))
                {
                    Console.WriteLine("The checkin file doesn't exist exist on disk.");
                    return "";
                }


                if (fileIteration.IsCheckedOut)
                {
                    VaultConn.ActiveConnection.FileManager.CheckinFile(fileIteration, comment, false, File_GetAssociations(file.Name),
                                                        null, false, null, ACW.FileClassification.None, false, filePathAbs);
                }

                return filePathAbs.FileName;

            }

            return "";


        }

        public static VDF.Currency.FilePathAbsolute? GetFilePathByFile(string fileName)
        {
            if (VaultConn.ActiveConnection != null)
            {
                ACW.File file = VaultFileUtilities.File_FindByFileName(fileName);

                string workingFolder = VaultConn.ActiveConnection.WorkingFoldersManager.GetWorkingFolder("$").FullPath;

                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;

                ACW.Folder[] folders = docService.FindFoldersByIds(new long[] { file.FolderId });

                VDF.Currency.FilePathAbsolute filePathAbs = new VDF.Currency.FilePathAbsolute(folders[0].FullName);


                string updatedfolderPath = filePathAbs.FullPath.Replace("/", @"\");

                string folderPathWithWorking = updatedfolderPath.Replace("$\\", workingFolder);

                string filePathWithWorking = folderPathWithWorking + @"\" + file.Name;

                filePathAbs = new VDF.Currency.FilePathAbsolute(filePathWithWorking);

                return filePathAbs;


            }

            return null;


        }


        public static string File_GetLatest(string fileName)
        {
            // THIS FUNCTION "GETS" THE LATEST VERSION OF A FILE THAT IS AT A SPECIFIED LIFECYCLE STATE
            try
            {

                // Check Vault connection
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                // Get Vault services
                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;

                // Find file by name
                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new InvalidOperationException($"File '{fileName}' not found.");
                }

                // Get file iteration and full file path
                FileIteration fileIteration = new FileIteration(VaultConn.ActiveConnection, file);
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                // Acquire the file from Vault
                File_Acquire(file, false);

                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                // Handle invalid arguments
                Console.WriteLine($"File_GetLatestLifecycleStateVersion Failed: {ex.Message}");
                return "";
            }
        }
    }
}
