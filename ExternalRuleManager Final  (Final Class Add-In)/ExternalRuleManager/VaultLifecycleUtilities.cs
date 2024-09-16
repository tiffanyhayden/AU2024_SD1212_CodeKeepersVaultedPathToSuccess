using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;

namespace ExternalRuleManager
{
    internal class VaultLifecycleUtilities
    {

        public static string GetLifecycleState(string strFileName)
        {
            try
            {
                // Check if Vault connection is active
                if (VaultConn.ActiveConnection == null)
                {
                    throw new InvalidOperationException("Vault connection is not active.");
                }

                // Check if the fileName is valid
                if (string.IsNullOrEmpty(strFileName))
                {
                    throw new ArgumentException("File name must have a value to continue.", nameof(strFileName));
                }

                // Get the document service
                ACW.DocumentService docService = VaultConn.ActiveConnection.WebServiceManager.DocumentService;

                // Find the file by name
                ACW.File file = VaultFileUtilities.File_FindByFileName(strFileName);
                if (file == null)
                {
                    throw new FileNotFoundException($"File '{strFileName}' not found in Vault.");
                }

                // Get all versions of the file
                ACW.File[] fileVersions = docService.GetFilesByMasterId(file.MasterId);
                if (fileVersions == null || fileVersions.Length == 0)
                {
                    throw new InvalidOperationException($"No file versions found for '{strFileName}'.");
                }

                // Reverse the array to get the latest version first
                Array.Reverse(fileVersions);

                // Return the lifecycle state of the latest version
                return fileVersions[0].FileLfCyc.LfCycStateName;
            }
            catch (ArgumentException ex)
            {
                // Handle invalid file name input
                Console.WriteLine($"GetLifecycleState Failed: {ex.Message}");
                return "";
            }
        }

    }
}
