using ExternalRuleManager.Properties;
using Inventor;
using System.Linq.Expressions;

namespace ExternalRuleManager
{
    /// <summary>
    /// Manages the creation and customization of a custom ribbon in the Autodesk Inventor UI.
    /// </summary>
    public class CustomRibbon
    {
        #region Properties

        /// <summary>
        /// Gets or sets the button definition for the TestButton.
        /// </summary>
        private ButtonDefinition _testButton;
        public ButtonDefinition TestButton
        {
            get { return _testButton; }
            private set { _testButton = value; }
        }

        private ComboBoxDefinition _externalRules;
        public ComboBoxDefinition ExternalRules
        {
            get { return _externalRules; }
            private set { _externalRules = value; }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomRibbon"/> class.
        /// Creates the TestButton and adds it to the specified ribbons in the Inventor UI.
        /// </summary>
        public CustomRibbon()
        {
            Image image16x16 = ByteArrayToImage(Resources._16x16);
            Image image32x32 = ByteArrayToImage(Resources._32x32);

            TestButton = Utilities.CreateButtonDef("TestButton", "TestButton", "", image16x16, image32x32);
            TestButton.OnExecute += ExecuteTestButton;

            ExternalRules = Utilities.CreateComboBoxDef("External Rules", "ExternalRules", CommandTypesEnum.kQueryOnlyCmdType, 200);
            ExternalRules.AddItem("No External Rule Selected");
            ExternalRules.ListIndex = 1;


            Ribbon partRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            Ribbon assemblyRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            Ribbon drawingRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            Ribbon zeroDocRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["ZeroDoc"];

            AddToRibbon(partRibbon);
            AddToRibbon(assemblyRibbon);
            AddToRibbon(drawingRibbon);
            AddToRibbon(zeroDocRibbon);
        }

        #endregion

        /// <summary>
        /// Converts a byte array to an Image.
        /// </summary>
        /// <param name="byteArray">The byte array to convert.</param>
        /// <returns>The converted Image object.</returns>
        private Image ByteArrayToImage(byte[] byteArray)
        {
            using (MemoryStream ms = new MemoryStream(byteArray))
            {
                return Image.FromStream(ms);
            }
        }

        /// <summary>
        /// Adds the TestButton to a specific ribbon in the Inventor UI.
        /// </summary>
        /// <param name="targetRibbon">The ribbon to which the TestButton will be added.</param>
        private void AddToRibbon(Ribbon targetRibbon)
        {
            RibbonTab ermRibbonTab = targetRibbon.RibbonTabs.Add("External Rule Manager", "id_ExternalRuleManager" + targetRibbon.InternalName, Globals.InvAppGuidID);

            RibbonPanel ermRibbonPanel = ermRibbonTab.RibbonPanels.Add("Manage", "id_Manage" + targetRibbon.InternalName, Globals.InvAppGuidID);


            ermRibbonPanel.CommandControls.AddComboBox(ExternalRules);

            if (targetRibbon.InternalName != "ZeroDoc")
            {
                if (targetRibbon.InternalName == "Part")
                {
                    // Additional logic for Part ribbon can be added here
                }
                else if (targetRibbon.InternalName == "Assembly")
                {
                    // Additional logic for Assembly ribbon can be added here
                }
                else if (targetRibbon.InternalName == "Drawing")
                {
                    // Additional logic for Drawing ribbon can be added here
                }
            }
        }

        #region Ui Events

        /// <summary>
        /// Handles the execution event when the TestButton is clicked.
        /// Displays a message box indicating that the button has been pushed.
        /// </summary>
        /// <param name="context">The context in which the button was executed.</param>
        private void ExecuteTestButton(NameValueMap context)
        {

            try
            {
                MessageBox.Show(VaultConn.ActiveConnection.UserName, "External Rule Manager");
            }
            catch (Exception)
            {

                MessageBox.Show("Please Login To Vault through Inventor...", "External Rule Manager");
            }
            
            
        }

        #endregion
    }
}