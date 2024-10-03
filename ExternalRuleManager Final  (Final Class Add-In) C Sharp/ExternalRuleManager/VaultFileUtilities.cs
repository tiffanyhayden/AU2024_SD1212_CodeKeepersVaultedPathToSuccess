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

namespace ExternalRuleManager
{
    /// <summary>
    /// Provides utility methods for interacting with files in Autodesk Vault.
    /// </summary>
    internal class VaultFileUtilities
    {
        /// <summary>
        /// Acquires a file from Vault, optionally checking it out.
        /// </summary>
        /// <param name="file">The Vault file to acquire.</param>
        /// <param name="doCheckOut">If true, the file will be checked out; otherwise, it will just be downloaded.</param>
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

                // Set relationship gathering settings
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = false;
                oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = false;

                if (doCheckOut)
                {
                    oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;
                }
                else
                {
                    oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }

                oSettings.AddFileToAcquire(oFileIteration, oSettings.DefaultAcquisitionOption);
                VaultConn.ActiveConnection.FileManager.AcquireFiles(oSettings);
            }
        }

        /// <summary>
        /// Checks out a file from Vault.
        /// </summary>
        /// <param name="file">The Vault file to check out.</param>
        public static void File_CheckOut(ACW.File file)
        {
            IntPtr parent = IntPtr.Zero;
            InteractiveAcquireFileSettings oSettings;
            FileIteration oFileIteration;

            if (VaultConn.ActiveConnection != null)
            {
                try
                {
                    oFileIteration = new FileIteration(VaultConn.ActiveConnection, file);

                    if (oFileIteration.IsCheckedOut)
                    {
                        throw new ArgumentException("File is already checked out.");
                    }

                    oSettings = new InteractiveAcquireFileSettings(VaultConn.ActiveConnection, parent, "Download files");
                    oSettings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                    oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout;

                    oSettings.AddFileToAcquire(oFileIteration, oSettings.DefaultAcquisitionOption);
                    VaultConn.ActiveConnection.FileManager.AcquireFiles(oSettings);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error checking out the file.", ex);
                }
            }
        }

        /// <summary>
        /// Retrieves the latest version of a file in Vault that is at a specified lifecycle state.
        /// </summary>
        /// <param name="fileName">The name of the file to retrieve.</param>
        /// <param name="lifecycleStateName">The lifecycle state of the file (e.g., "Released").</param>
        /// <returns>The full path to the latest file version.</returns>
        public static string File_GetLatestLifecycleStateVersion(string fileName, string lifecycleStateName)
        {
            try
            {
                lifecycleStateName = lifecycleStateName.ToUpper();

                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(lifecycleStateName))
                {
                    throw new ArgumentException("File name or lifecycle state cannot be null or empty.");
                }

                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;

                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new InvalidOperationException($"File '{fileName}' not found.");
                }

                ACW.File[] fileVersions = docService.GetFilesByMasterId(file.MasterId);
                if (fileVersions == null || fileVersions.Length == 0)
                {
                    throw new InvalidOperationException($"No versions found for file '{fileName}'.");
                }

                Array.Reverse(fileVersions);

                ACW.File fileByLifecycleState = fileVersions.FirstOrDefault(fileVersion => fileVersion.FileLfCyc.LfCycStateName.ToUpper() == lifecycleStateName);

                if (fileByLifecycleState == null)
                {
                    throw new InvalidOperationException($"No file found with lifecycle state '{lifecycleStateName}'.");
                }

                FileIteration fileIteration = new FileIteration(VaultConn.ActiveConnection, fileByLifecycleState);
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                File_Acquire(fileByLifecycleState, false);

                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"File_GetLatestLifecycleStateVersion Failed: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Finds a file in Vault by its name.
        /// </summary>
        /// <param name="fileName">The name of the file to search for.</param>
        /// <param name="srchOperator">The search operator to use (default is 3).</param>
        /// <param name="searchInStartingFolder">The folder in which to start the search (default is the root "$/").</param>
        /// <returns>The first file that matches the search criteria.</returns>
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

                ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.SysName == "ClientFileName");
                if (propDef == null)
                {
                    throw new InvalidOperationException("Property definition for 'ClientFileName' not found.");
                }

                ACW.SrchCond search = new ACW.SrchCond
                {
                    PropDefId = propDef.Id,
                    PropTyp = ACW.PropertySearchType.SingleProperty,
                    SrchOper = srchOperator,
                    SrchRule = ACW.SearchRuleType.Must,
                    SrchTxt = fileName
                };

                ACW.Folder[] folders = docService.FindFoldersByPaths(new string[] { searchInStartingFolder });
                if (folders == null || folders.Length == 0)
                {
                    throw new InvalidOperationException($"No folders found at the specified path: {searchInStartingFolder}");
                }

                long[] folderIDs = folders.Where(f => f.Id != -1).Select(f => f.Id).ToArray();
                if (folderIDs == null || folderIDs.Length == 0)
                {
                    throw new InvalidOperationException("No valid folder IDs found.");
                }

                string bookmark = string.Empty;
                ACW.SrchStatus status;
                ACW.File[] results = docService.FindFilesBySearchConditions(new ACW.SrchCond[] { search }, null, folderIDs, true, true, ref bookmark, out status);

                if (results == null || results.Length == 0)
                {
                    throw new InvalidOperationException($"No files found matching the search criteria: {fileName}");
                }

                return results[0];
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine("File_FindByFileName Failed: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Undoes the checkout of a file in Vault.
        /// </summary>
        /// <param name="fileName">The name of the file to undo the checkout for.</param>
        /// <returns>The full path of the file after undoing the checkout.</returns>
        public static string File_UndoCheckOut(string fileName)
        {
            try
            {
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("File name must have a value to continue.", nameof(fileName));
                }

                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new FileNotFoundException($"File '{fileName}' not found in Vault.");
                }

                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                if (File_IsCheckedOut(file.Name))
                {
                    VaultConn.ActiveConnection.FileManager.UndoCheckoutFile(fileIteration);
                }

                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"File_UndoCheckOut Failed: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Checks if a file is currently checked out in Vault.
        /// </summary>
        /// <param name="fileName">The name of the file to check.</param>
        /// <param name="toCurrentUser">If true, only check if the file is checked out to the current user.</param>
        /// <returns>True if the file is checked out; otherwise, false.</returns>
        public static bool File_IsCheckedOut(string fileName, bool toCurrentUser = false)
        {
            try
            {
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    throw new ArgumentException("File name must have a value to continue.", nameof(fileName));
                }

                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new FileNotFoundException($"File '{fileName}' was not found in Vault.");
                }

                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                return !toCurrentUser ? fileIteration.IsCheckedOut : fileIteration.IsCheckedOutToCurrentUser;
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"File_IsCheckedOut Failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Retrieves all file associations for the specified file in Vault.
        /// </summary>
        /// <param name="fileName">The name of the file to get associations for.</param>
        /// <returns>An array of <see cref="FileAssocParam"/> objects representing the file associations.</returns>
        public static FileAssocParam[] File_GetAssociations(string fileName)
        {
            if (VaultConn.ActiveConnection != null)
            {
                ACW.File file = File_FindByFileName(fileName);
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                VDF.Vault.Settings.FileRelationshipGatheringSettings relationshipSettings = new VDF.Vault.Settings.FileRelationshipGatheringSettings
                {
                    IncludeAttachments = true,
                    IncludeChildren = true,
                    IncludeParents = true,
                    IncludeRelatedDocumentation = true
                };

                IEnumerable<ACW.FileAssocLite> fileAssocLite = VaultConn.ActiveConnection.FileManager.GetFileAssociationLites(
                    new long[] { fileIteration.EntityIterationId },
                    relationshipSettings
                );

                List<FileAssocParam> fileAssocParams = new List<FileAssocParam>();

                if (fileAssocLite != null)
                {
                    foreach (ACW.FileAssocLite thisFileAssocLite in fileAssocLite)
                    {
                        if (thisFileAssocLite.CldFileId != fileIteration.EntityIterationId)
                        {
                            FileAssocParam par = new FileAssocParam
                            {
                                CldFileId = thisFileAssocLite.CldFileId,
                                RefId = thisFileAssocLite.RefId,
                                Source = thisFileAssocLite.Source,
                                Typ = thisFileAssocLite.Typ,
                                ExpectedVaultPath = thisFileAssocLite.ExpectedVaultPath
                            };

                            fileAssocParams.Add(par);
                        }
                    }
                }

                return fileAssocParams.ToArray();
            }

            return null;
        }

        /// <summary>
        /// Checks in a file to Vault with the specified comment.
        /// </summary>
        /// <param name="fileName">The name of the file to check in.</param>
        /// <param name="comment">The comment to associate with the check-in.</param>
        /// <returns>The name of the file after checking it in.</returns>
        public static string File_CheckIn(string fileName, string comment)
        {
            if (VaultConn.ActiveConnection != null)
            {
                ACW.File file = File_FindByFileName(fileName);

                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(VaultConn.ActiveConnection, file);

                if (!fileIteration.IsCheckedOut)
                {
                    throw new FileNotFoundException("File is not checked out.");
                }

                VDF.Currency.FilePathAbsolute? filePathAbs = GetFilePathByFile(fileName);

                if (!System.IO.File.Exists(filePathAbs.FullPath))
                {
                    Console.WriteLine("The check-in file doesn't exist on disk.");
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

        /// <summary>
        /// Gets the file path on disk for a specified Vault file.
        /// </summary>
        /// <param name="fileName">The name of the Vault file.</param>
        /// <returns>A <see cref="VDF.Currency.FilePathAbsolute"/> object representing the file path on disk, or <c>null</c> if the file path cannot be retrieved.</returns>
        public static VDF.Currency.FilePathAbsolute? GetFilePathByFile(string fileName)
        {
            if (VaultConn.ActiveConnection != null)
            {
                ACW.File file = File_FindByFileName(fileName);

                string workingFolder = VaultConn.ActiveConnection.WorkingFoldersManager.GetWorkingFolder("$").FullPath;

                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                ACW.Folder[] folders = docService.FindFoldersByIds(new long[] { file.FolderId });

                VDF.Currency.FilePathAbsolute filePathAbs = new VDF.Currency.FilePathAbsolute(folders[0].FullName);
                string updatedFolderPath = filePathAbs.FullPath.Replace("/", @"\");
                string folderPathWithWorking = updatedFolderPath.Replace("$\\", workingFolder);
                string filePathWithWorking = folderPathWithWorking + @"\" + file.Name;

                filePathAbs = new VDF.Currency.FilePathAbsolute(filePathWithWorking);
                return filePathAbs;
            }

            return null;
        }

        /// <summary>
        /// Retrieves the latest version of a file from Vault.
        /// </summary>
        /// <param name="fileName">The name of the file to retrieve.</param>
        /// <returns>The full path of the latest version of the file.</returns>
        public static string File_GetLatest(string fileName)
        {
            try
            {
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;
                VDF.Vault.Services.Connection.IWorkingFoldersManager services = VaultConn.ActiveConnection.WorkingFoldersManager;

                ACW.File file = File_FindByFileName(fileName);
                if (file == null)
                {
                    throw new InvalidOperationException($"File '{fileName}' not found.");
                }

                FileIteration fileIteration = new FileIteration(VaultConn.ActiveConnection, file);
                string fullFileName = services.GetPathOfFileInWorkingFolder(fileIteration).FullPath;

                File_Acquire(file, false);

                return fullFileName;
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"File_GetLatest Failed: {ex.Message}");
                return "";
            }
        }
    }
}
