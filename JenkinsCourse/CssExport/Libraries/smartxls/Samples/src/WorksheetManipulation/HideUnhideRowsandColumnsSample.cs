using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class HideUnhideRowsandColumnsSample : BaseFileViewControl, ISampleControl
    {
        public HideUnhideRowsandColumnsSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = new WorkBook();
            try
            {
                m_book.read("..\\template\\book.xls");
                m_book.setRowHidden(1,true);
                m_book.setColHidden(1,true);
                m_book.write(".\\Sample.xls");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("Sample.xls");
            }
            else
            {
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\WorksheetManipulation\\HideUnhideRowsandColumnsSample.cs";
            }
        }
    }
}
