﻿#pragma checksum "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "5EF793E036199F5971C658B23D6CA828D5E360308E9635FFEA4673E1481C72F4"
//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using Магазин_техники.Pages;


namespace Магазин_техники.Pages {
    
    
    /// <summary>
    /// PageAddInfornationAboutTecnica
    /// </summary>
    public partial class PageAddInfornationAboutTecnica : System.Windows.Controls.Page, System.Windows.Markup.IComponentConnector {
        
        
        #line 36 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox cbTypetechnica;
        
        #line default
        #line hidden
        
        
        #line 48 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox tbxType;
        
        #line default
        #line hidden
        
        
        #line 53 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnAddType;
        
        #line default
        #line hidden
        
        
        #line 65 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox cbManuf;
        
        #line default
        #line hidden
        
        
        #line 80 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox tbxManuf;
        
        #line default
        #line hidden
        
        
        #line 87 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnAddManuf;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/WpfApp1;component/pages/pageaddinfornationabouttecnica.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.cbTypetechnica = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 2:
            this.tbxType = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            this.btnAddType = ((System.Windows.Controls.Button)(target));
            
            #line 54 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
            this.btnAddType.Click += new System.Windows.RoutedEventHandler(this.btnAddType_Click);
            
            #line default
            #line hidden
            return;
            case 4:
            this.cbManuf = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 5:
            this.tbxManuf = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.btnAddManuf = ((System.Windows.Controls.Button)(target));
            
            #line 88 "..\..\..\Pages\PageAddInfornationAboutTecnica.xaml"
            this.btnAddManuf.Click += new System.Windows.RoutedEventHandler(this.btnAddManuf_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}
