﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.36468
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Report_pack_generator.Settings {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0")]
    internal sealed partial class color_settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static color_settings defaultInstance = ((color_settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new color_settings())));
        
        public static color_settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("0, 192, 192")]
        public global::System.Drawing.Color inprogress {
            get {
                return ((global::System.Drawing.Color)(this["inprogress"]));
            }
            set {
                this["inprogress"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("LimeGreen")]
        public global::System.Drawing.Color completed {
            get {
                return ((global::System.Drawing.Color)(this["completed"]));
            }
            set {
                this["completed"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Default")]
        public string theme {
            get {
                return ((string)(this["theme"]));
            }
            set {
                this["theme"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("LightSlateGray")]
        public global::System.Drawing.Color colorscheme {
            get {
                return ((global::System.Drawing.Color)(this["colorscheme"]));
            }
            set {
                this["colorscheme"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool ThemeChanged {
            get {
                return ((bool)(this["ThemeChanged"]));
            }
            set {
                this["ThemeChanged"] = value;
            }
        }
    }
}
