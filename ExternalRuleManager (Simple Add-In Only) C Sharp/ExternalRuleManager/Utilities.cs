using Inventor;
using static ExternalRuleManager.ImageConverter;

namespace ExternalRuleManager
{
    public class Utilities


    {
        public static Inventor.ButtonDefinition CreateButtonDef(string DisplayName,
            string InternalName,
            string ToolTip = "",
            Image SmallIcon = null,
            Image LargeIcon = null)
        {

            Inventor.ButtonDefinition? existingDef = null;

            try
            {
                existingDef = Globals.InvApp.CommandManager.ControlDefinitions[InternalName] as ButtonDefinition;
            }
            catch (Exception)
            {

            }

            if (existingDef != null)

            {
                //Add message box for user maybe? 
                return null;
            }

            IPictureDisp iPic16 = null;
            if (SmallIcon != null)
            {
                try
                {
                    iPic16 = ConvertImageToIPictureDisp(SmallIcon);
                }
                catch (Exception)
                {

                    //Add msgbox for user maybe? 
                }
            }

            IPictureDisp iPic32 = null;
            if (LargeIcon != null)
            {
                try
                {
                    iPic32 = ConvertImageToIPictureDisp(LargeIcon);
                }
                catch (Exception)
                {


                    //Add msgbox for user maybe? 
                }

            }

            try
            {
                Inventor.ControlDefinitions controlDefs = Globals.InvApp.CommandManager.ControlDefinitions;

                ButtonDefinition btnDef = controlDefs.AddButtonDefinition(DisplayName,
                                                                            InternalName,
                                                                            CommandTypesEnum.kShapeEditCmdType,
                                                                            Globals.InvAppGuidID,
                                                                            "",
                                                                            ToolTip,
                                                                            iPic16,
                                                                            iPic32);
                return btnDef;



            }
            catch (Exception)
            {

                return null;
            }


        }
    }

}







