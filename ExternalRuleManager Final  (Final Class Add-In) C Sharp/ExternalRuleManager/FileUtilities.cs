using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;
using File = System.IO.File;


namespace ExternalRuleManager
{
    internal class FileUtilities
    {

        private const string TabName = "External Rule Manager";



        /// <summary>
        /// Makes a local copy of a selected external rule file from the Vault.
        /// </summary>
        /// <remarks>
        /// The copied file will be renamed to include the current user's username. 
        /// If the file already exists at the destination, a message is shown and no copy is made.
        /// </remarks>
        public static void MakeLocalCopy()
        {
            try
            {
                // Get the current username
                string username = System.Environment.UserName;

                // Ensure a valid item is selected
                if (CustomRibbon.ExternalRules.ListIndex <= 0 || CustomRibbon.ExternalRules.ListIndex == 1)
                {
                    Console.WriteLine("No valid item selected in the ExternalRules combo box.");
                    return;
                }

                string selectedItemName = CustomRibbon.ExternalRules.ListItem[CustomRibbon.ExternalRules.ListIndex];

                // Get the file path and check if the file exists
                VDF.Currency.FilePathAbsolute? filePathAbs = VaultFileUtilities.GetFilePathByFile(selectedItemName);
                if (filePathAbs == null || string.IsNullOrEmpty(filePathAbs.FullPath))
                {
                    Console.WriteLine($"Could not find a local file path for '{selectedItemName}'.");
                    return;
                }

                //Get the file
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                if (File.Exists(filePathAbs.FullPath))
                {
                    string sourceFile = filePathAbs.FullPath;
                    string pathWithoutExt = System.IO.Path.GetFileNameWithoutExtension(sourceFile);
                    string destFile = $"{pathWithoutExt}_{username}{System.IO.Path.GetExtension(sourceFile)}";
                    string folderPath = filePathAbs.FolderPath;
                    string destPath = System.IO.Path.Combine(folderPath, destFile);

                    // Ensure destination folder is valid
                    if (!Directory.Exists(folderPath))
                    {
                        Console.WriteLine($"Destination folder '{folderPath}' does not exist.");
                        return;
                    }

                    // Try to copy the file
                    try
                    {
                        if (!File.Exists(destPath))
                        {
                            File.Copy(sourceFile, destPath);
                            FileInfo fileInfo = new FileInfo(destPath);
                            if (fileInfo.IsReadOnly)
                            {
                                fileInfo.IsReadOnly = false;
                            }

                            Console.WriteLine($"File copied successfully to {destPath}");
                        }
                        else
                        {
                            MessageBox.Show($"Cannot make copy, file already exists at destination: {destPath}", TabName);
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"I/O error during file copy: {ex.Message}", TabName);
                    }
                }
                else
                {
                    MessageBox.Show($"Source file '{filePathAbs.FullPath}' does not exist on disk.", TabName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", TabName);
            }
        }

        /// <summary>
        /// Overwrites an external rule on disk with a local copy.
        /// </summary>
        /// <remarks>
        /// If the file is not checked out, the user is prompted for confirmation before overwriting.
        /// </remarks>
        public static void OverwriteRuleOnDiskWithCopy()
        {
            try
            {
                // Get the current username
                string username = System.Environment.UserName;

                // Ensure a valid item is selected
                if (CustomRibbon.ExternalRules.ListIndex <= 0 || CustomRibbon.ExternalRules.ListIndex == 1)
                {
                    Console.WriteLine("No valid item selected in the ExternalRules combo box.");
                    return;
                }

                //Get the name of the selected item from the combobox
                string selectedItemName = CustomRibbon.ExternalRules.ListItem[CustomRibbon.ExternalRules.ListIndex];

                // Get the file path and check if the file exists
                VDF.Currency.FilePathAbsolute? filePathAbs = VaultFileUtilities.GetFilePathByFile(selectedItemName);
                if (filePathAbs == null || string.IsNullOrEmpty(filePathAbs.FullPath))
                {
                    Console.WriteLine($"Could not find a local file path for '{selectedItemName}'.");
                    return;
                }

                //Get the file
                ACW.File file = VaultFileUtilities.File_FindByFileName(selectedItemName);

                if (File.Exists(filePathAbs.FullPath))
                {
                    string sourceFile = filePathAbs.FullPath;
                    string pathWithoutExt = System.IO.Path.GetFileNameWithoutExtension(sourceFile);
                    string destFile = $"{pathWithoutExt}_{username}{Path.GetExtension(sourceFile)}";
                    string folderPath = filePathAbs.FolderPath;
                    string destPath = System.IO.Path.Combine(folderPath, destFile);


                    // Ensure destination folder is valid
                    if (!Directory.Exists(folderPath))
                    {
                        Console.WriteLine($"Destination folder '{folderPath}' does not exist.");
                        return;
                    }

                    // Try to copy the file
                    try
                    {
                        if (File.Exists(destPath))
                        {
                            if (VaultFileUtilities.File_IsCheckedOut(file.Name))
                            {
                                ReplaceFile(destPath, sourceFile);
                            }
                            else
                            {
                                DialogResult result = MessageBox.Show($"Vaulted version is not checked out, do you still wish to overwrite?", TabName, MessageBoxButtons.YesNo);

                                if (result == DialogResult.Yes)
                                {
                                    ReplaceFile(destPath, sourceFile);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show($"File does not exist on disk: {destPath}", TabName);
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"I/O error during file copy: {ex.Message}", TabName);
                    }
                }
                else
                {
                    MessageBox.Show($"Source file '{filePathAbs.FullPath}' does not exist on disk.", TabName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", TabName);
            }
        }

        /// <summary>
        /// Replaces the target file with the source file, ensuring the target file is writable.
        /// </summary>
        /// <param name="sourceFile">The path of the source file to replace with.</param>
        /// <param name="targetFile">The path of the target file to be replaced.</param>
        /// <remarks>
        /// Uses an atomic operation to ensure that the target file is safely replaced. If the target file is read-only, the attribute is removed before replacement.
        /// </remarks>
        private static void ReplaceFile(string sourceFile, string targetFile)
        {
            try
            {
                FileInfo targetFileInfo = new FileInfo(targetFile);

                // Ensure the target file is not read-only
                if (targetFileInfo.IsReadOnly)
                {
                    targetFileInfo.IsReadOnly = false;
                }

               
                // Use File.Replace
                File.Replace(sourceFile, targetFile, null); 
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show($"Target file not found: {ex.Message}", TabName);
            }
            catch (IOException ex)
            {
                MessageBox.Show($"I/O error during file replace: {ex.Message}", TabName);
            }
        }
    }
}