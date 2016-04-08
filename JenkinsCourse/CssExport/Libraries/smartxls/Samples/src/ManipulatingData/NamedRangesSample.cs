using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class NamedRangesSample : BaseFileViewControl, ISampleControl
    {
        public NamedRangesSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook book = new WorkBook();
            book.read("..\\template\\NamedRanges.xls");

            book.setDefinedName("Products", "$A$1:$A$6");
         
            book.setDefinedName("One", "$C$3");
            book.setDefinedName("Two", "$D$3");
            book.setSelection("E3");
            book.setFormula(2, 4, "SUM(One, Two)");
            book.recalc();

            //Saving the workbook to disk.
            book.write(".\\Sample.xls");

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("Sample.xls");
            }
            else
            {}
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\ManipulatingData\\NamedRangesSample.cs";
            }
        }
    }
}
