using System.Diagnostics;
using ACW = Autodesk.Connectivity.WebServices;


namespace ExternalRuleManager
{
    internal class VaultUtilities
    {
        /// <summary>
        /// Retrieves the latest files from a specified folder in the Vault.
        /// </summary>
        /// <param name="folderName">The name of the folder to retrieve files from.</param>
        /// <returns>
        /// An array of <see cref="ACW.File"/> representing the files in the folder, 
        /// or <c>null</c> if the folder is not found or if the Vault connection is not established.
        /// </returns>
        public static ACW.File[]? GetFilesByFolder(string folderName)
        {
            if (VaultConn.IsConnected())
            {
                ACW.Folder[]? folders = VaultFolderUtilities.GetFoldersByFolderName(folderName);
                ACW.Folder? folder = folders[0];
                Autodesk.Connectivity.WebServicesTools.WebServiceManager? webServiceManager = VaultConn.ActiveConnection?.WebServiceManager;
                ACW.DocumentService? docService = webServiceManager.DocumentService;
                ACW.File[]? files = docService.GetLatestFilesByFolderId(folder.Id, false);

                // If files are found in the main folder, return them
                if (files != null && files.Length > 0)
                {
                    return files;
                }


                long[] subfolderIds = docService.GetFolderIdsByParentIds(new long[] { (long)folder.Id }, true);


                List<ACW.File> allFiles = new List<ACW.File>();

                foreach (long subfolderId in subfolderIds)
                {
                    // Retrieve files in each subfolder
                    ACW.File[]? subfolderFiles = docService.GetLatestFilesByFolderId(subfolderId, false);
                    if (subfolderFiles != null)
                    {
                        allFiles.AddRange(subfolderFiles);
                    }
                }

                // Return files from subdirectories if any were found, otherwise return null
                return allFiles.Count > 0 ? allFiles.ToArray() : null;

            }

            return null;
        }

        /// <summary>
        /// Retrieves the latest folder with a specified name from the Vault.
        /// </summary>
        /// <param name="folderName">The name of the folder to retrieve.</param>
        /// <returns>
        /// The latest <see cref="ACW.Folder"/> object found in the Vault or <c>null</c> if no folder is found.
        /// </returns>
        public static ACW.Folder? GetLatestOnFolder(string folderName)
        {
            ACW.Folder? folder = null;
            ACW.Folder[]? folders = null;

            try
            {
                folders = VaultFolderUtilities.GetFoldersByFolderName("EXTERNAL RULES");

                if (folders != null && folders.Length > 0)
                {
                    folder = folders[0];
                }

                if (folder != null)
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

        /// <summary>
        /// Retrieves the latest version of files within a specified folder based on the lifecycle state.
        /// </summary>
        /// <param name="folderName">The name of the folder to retrieve files from.</param>
        /// <param name="lifecycleState">The lifecycle state to filter the files by.</param>
        /// <remarks>
        /// This method logs errors to the console for files that cannot be processed or for validation issues.
        /// </remarks>
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
