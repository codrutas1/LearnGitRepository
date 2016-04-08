using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace SmartXLSExplorer
{
    /// <summary>
    /// This class is only used to manage the SmartXLS Explorer.  
    /// Look at the separate sample user controls for sample source code.
    /// These user controls are found under the folders matching the name
    /// of the sample category such as the IRange or WorkbookView folder.
    /// The relevant sample code is in the *.code.cs files matching the
    /// name of each sample.
    /// </summary>
    public partial class SampleNavigationControl : UserControl
    {
        [Category("Action"),
         Description("Occurs when the selection changes.")]
        public event SelectionChangedEventHandler SelectionChanged;

        public SampleNavigationControl()
        {
            InitializeComponent();
            treeView1.ExpandAll();
        }

        protected virtual void OnSelectionChanged(TreeViewEventArgs e)
        {
            if (SelectionChanged != null)
                SelectionChanged(this, e);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (SelectionChanged != null)
                OnSelectionChanged(e);
        }
    }

    public delegate void SelectionChangedEventHandler(object sender, TreeViewEventArgs e);

}
