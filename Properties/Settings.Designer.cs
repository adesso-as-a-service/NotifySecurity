﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace NotifySecurity.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.2.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string Security_Team_Mail_cc {
            get {
                return ((string)(this["Security_Team_Mail_cc"]));
            }
            set {
                this["Security_Team_Mail_cc"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string Security_Team_Mail_bcc {
            get {
                return ((string)(this["Security_Team_Mail_bcc"]));
            }
            set {
                this["Security_Team_Mail_bcc"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://swordphish.domain.com/")]
        public string SwordphishURL {
            get {
                return ((string)(this["SwordphishURL"]));
            }
            set {
                this["SwordphishURL"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("X-Swordphish-Awareness-Campaign")]
        public string SwordPhishHeader {
            get {
                return ((string)(this["SwordPhishHeader"]));
            }
            set {
                this["SwordPhishHeader"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("security@adesso-service.com")]
        public string Security_Team_Mail {
            get {
                return ((string)(this["Security_Team_Mail"]));
            }
        }



        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("security@smarthouse.de")]
        public string Security_Team_Mail1
        {
            get
            {
                return ((string)(this["Security_Team_Mail1"]));
            }
        }
    }
}
