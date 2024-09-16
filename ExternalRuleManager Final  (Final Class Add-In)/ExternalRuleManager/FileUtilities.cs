using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VDF = Autodesk.DataManagement.Client.Framework;
using ACW = Autodesk.Connectivity.WebServices;
using File = System.IO.File;

namespace ExternalRuleManager
{
    internal class FileUtilities
    {

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
                            Console.WriteLine($"File copied successfully to {destPath}");
                        }
                        else
                        {
                            MessageBox.Show($"Cannot make copy, file already exists at destination: {destPath}", "External Rule Manager");
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"I/O error during file copy: {ex.Message}", "External Rule Manager");
                    }
                }
                else
                {
                    MessageBox.Show($"Source file '{filePathAbs.FullPath}' does not exist.", "External Rule Manager");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "External Rule Manager");
            }
        }

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

                string selectedItemName = CustomRibbon.ExternalRules.ListItem[CustomRibbon.ExternalRules.ListIndex];

                // Get the file path and check if the file exists
                VDF.Currency.FilePathAbsolute? filePathAbs = VaultFileUtilities.GetFilePathByFile(selectedItemName);
                if (filePathAbs == null || string.IsNullOrEmpty(filePathAbs.FullPath))
                {
                    Console.WriteLine($"Could not find a local file path for '{selectedItemName}'.");
                    return;
                }

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
                            if(VaultFileUtilities.File_IsCheckedOut(file.Name))
                            {
                                File.Delete(sourceFile);
                                File.Move(destFile, sourceFile);

                            }
                            else
                            {
                                DialogResult result = MessageBox.Show($"Vaulted version is not checked out, do you still wish to overwrite?", "External Rule Manager", MessageBoxButtons.YesNo);

                                if(result == DialogResult.Yes)
                                {
                                    File.Delete(sourceFile);
                                    File.Move(destFile, sourceFile);

                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show($"File does not exist on disk: {destPath}", "External Rule Manager");
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"I/O error during file copy: {ex.Message}", "External Rule Manager");
                    }
                }
                else
                {
                    MessageBox.Show($"Source file '{filePathAbs.FullPath}' does not exist.", "External Rule Manager");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "External Rule Manager");
            }
        }

    }
}
