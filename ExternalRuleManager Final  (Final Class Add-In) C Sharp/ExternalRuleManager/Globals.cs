﻿namespace ExternalRuleManager
{
    /// <summary>
    /// A static class that contains global variables and constants for the Inventor add-in.
    /// </summary>
    public static class Globals
    {
        /// <summary>
        /// A private field that stores the instance of the Inventor application.
        /// </summary>
        private static Inventor.Application? _invApp;

        /// <summary>
        /// A private field that stores the event handler for the Inventor application events.
        /// </summary>
        private static InventorEventHandler? _invAppEvents;

        /// <summary>
        /// A private field that stores the event handler for the Vault application events.
        /// </summary>
        private static VaultEventHandler? _vaultAppEvents;

        /// <summary>
        /// A private field that stores the custom ribbon for the Inventor application.
        /// </summary>
        private static CustomRibbon? _invAppRibbon;

        /// <summary>
        /// Gets or sets the instance of the Inventor application. 
        /// This should be initialized before use.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown when trying to access the property if it has not been initialized.
        /// </exception>
        public static Inventor.Application InvApp
        {
            get
            {
                if (_invApp == null)
                {
                    // Handle initialization if required, or throw an exception
                    throw new InvalidOperationException("InvApp has not been initialized.");
                }
                return _invApp;
            }
            set
            {
                _invApp = value;
            }
        }

        /// <summary>
        /// Gets or sets the event handler for the Inventor application events.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown when trying to access the property if it has not been initialized.
        /// </exception>
        public static InventorEventHandler InvAppEvents
        {
            get
            {
                if (_invAppEvents == null)
                {
                    throw new InvalidOperationException("InvAppEvents has not been initialized.");
                }
                return _invAppEvents;
            }
            set
            {
                _invAppEvents = value;
            }
        }

        /// <summary>
        /// Gets or sets the event handler for the Vault application events.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown when trying to access the property if it has not been initialized.
        /// </exception>
        public static VaultEventHandler VaultAppEvents
        {
            get
            {
                if (_vaultAppEvents == null)
                {
                    throw new InvalidOperationException("VaultAppEvents has not been initialized.");
                }
                return _vaultAppEvents;
            }
            set
            {
                _vaultAppEvents = value;
            }
        }

        /// <summary>
        /// Gets or sets the custom ribbon for the Inventor application.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown when trying to access the property if it has not been initialized.
        /// </exception>
        public static CustomRibbon InvAppRibbon
        {
            get
            {
                if (_invAppRibbon == null)
                {
                    throw new InvalidOperationException("InvAppRibbon has not been initialized.");
                }
                return _invAppRibbon;
            }
            set
            {
                _invAppRibbon = value;
            }
        }

        /// <summary>
        /// A constant string that stores the GUID of the Inventor application.
        /// </summary>
        public const string InvAppGuid = "472DC0BC-8D55-4A5F-99AC-56DD17F66771";

        /// <summary>
        /// A constant string that stores the GUID of the Inventor application enclosed in curly braces.
        /// </summary>
        public const string InvAppGuidID = "{" + InvAppGuid + "}";

        /// <summary>
        /// A constant string that stores the name of the external rules directory.
        /// </summary>
        public const string ExternalRuleName = "EXTERNAL RULES";

        /// <summary>
        /// A constant string that stores the path of the external rules directory.
        /// </summary>
        public static string ExternalRuleDir = "C:\\AU2024\\AU\\EXTERNAL RULES";




    }
}