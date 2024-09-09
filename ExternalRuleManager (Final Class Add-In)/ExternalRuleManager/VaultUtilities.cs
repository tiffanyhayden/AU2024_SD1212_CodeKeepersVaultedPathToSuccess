using Autodesk.Connectivity.WebServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACW = Autodesk.Connectivity.WebServices;

namespace ExternalRuleManager
{
    internal class VaultUtilities
    {

        public ACW.Folder[]? GetFoldersByFolderName(string folderName)
        {
            Autodesk.Connectivity.WebServicesTools.WebServiceManager webServiceManager;
            ACW.DocumentService docService;
            ACW.Folder[] folders = null;
            ACW.PropDef[] propDefs = null;
            ACW.PropDef? propDef  = null;    
            ACW.SrchCond searchCondition = new ACW.SrchCond();
            string bookmark = "";
            ACW.SrchStatus status;


            if (VaultConn.ActiveConnection == null) 
                {
                    return null;
                }

            webServiceManager = VaultConn.ActiveConnection.WebServiceManager;
            docService = webServiceManager.DocumentService;
            propDefs = webServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR");
            propDef = Array.Find(propDefs, n => n.DispName == "Name");

            searchCondition.PropDefId = propDef.Id;
            searchCondition.PropTyp = ACW.PropertySearchType.SingleProperty;
            searchCondition.SrchOper = 3;
            searchCondition.SrchRule = ACW.SearchRuleType.Must;
            searchCondition.SrchTxt = folderName;



            try
            {
                folders = docService.FindFoldersBySearchConditions(new ACW.SrchCond[] { searchCondition }, null, new long[] { }, true, bookmark,  status);
            }
            catch (Exception)
            {
                // No folders found
                return null;
            }


            return folders;


        }








    }
}
