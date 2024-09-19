using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Forms.Settings;
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
    internal class VaultUtilities
    {




        public static ACW.File[]? GetFilesByFolder(string folderName)
        {

            if (VaultConn.IsConnected())
            {

                ACW.Folder[]? folders = VaultFolderUtilities.GetFoldersByFolderName(folderName);
                ACW.Folder? folder = folders[0];
                Autodesk.Connectivity.WebServicesTools.WebServiceManager? webServiceManager = VaultConn.ActiveConnection?.WebServiceManager;
                ACW.DocumentService? docService = webServiceManager.DocumentService;
                ACW.File[]? files = docService.GetLatestFilesByFolderId(folder.Id, false);

                return files;
            }


            return null;

        }


        public static ACW.Folder? GetLatestOnFolder(string folderName)
        {
            ACW.Folder? folder = null;
            ACW.Folder[]? folders = null;

            try
            {
                folders = VaultFolderUtilities.GetFoldersByFolderName("EXTERNAL RULES");
                
                if(folders != null && folders.Length > 0)
                {
                    folder = folders[0];
                }
                
                if(folder != null)
                {
                    VaultFolderUtilities.Folder_Acquire(folder);
                    return folder;
                }

                return null;
                
            }
            catch (Exception)
            {

                return null;
            }


        }

        public static void GetLatestFilesByLifecycleState(string folderName, string lifecycleState)
        {
            try
            {
                // Input validation
                if (string.IsNullOrEmpty(folderName))
                {
                    throw new ArgumentException("Folder name cannot be null or empty.", nameof(folderName));
                }

                if (string.IsNullOrEmpty(lifecycleState))
                {
                    throw new ArgumentException("Lifecycle state cannot be null or empty.", nameof(lifecycleState));
                }

                // Get the files from the specified folder
                ACW.File[]? files = GetFilesByFolder(folderName);

                if (files == null || files.Length == 0)
                {
                    Console.WriteLine($"No files found in folder: {folderName}");
                    return;
                }

                // Process each file
                foreach (ACW.File file in files)
                {
                    try
                    {
                        // Get the latest file by lifecycle state
                        VaultFileUtilities.File_GetLatestLifecycleStateVersion(file.Name, lifecycleState);
                    }
                    catch (Exception ex)
                    {
                        // Log specific errors for each file
                        Console.WriteLine($"Error processing file '{file.Name}': {ex.Message}");
                    }
                }
            }
            catch (ArgumentException ex)
            {
                // Handle argument validation errors
                Console.WriteLine($"Input error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }


        
    }
}
