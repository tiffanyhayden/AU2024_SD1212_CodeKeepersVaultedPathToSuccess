using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    internal class VaultFolderUtilities
    {

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


                //Check Connection before continuing. 
                if (VaultConn.IsConnected())
                {

                    try
                    {
                        //Create Folder Entity
                        folderEntity = new VDF.Vault.Currency.Entities.Folder(VaultConn.ActiveConnection, folder);
                    }
                    catch (Exception)
                    {

                        throw;
                    }




                    //*******************************************************************************************
                    // DEFINE SETTINGS
                    //*******************************************************************************************
                    settings = new VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings(VaultConn.ActiveConnection, parent, "Download files");
                    settings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                    settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                    settings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

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
                            MessageBox.Show("Acquiring file failed.", "Error", MessageBoxButtons.OK);
                            throw;
                        }


                    }

                }
            }

            return null;

        }

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
                searchCondition.SrchOper = 3;
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
