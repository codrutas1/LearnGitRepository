using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using SmartXLSExplorer;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class WriteImagesSample : BaseFileViewControl, ISampleControl
    {
        public WriteImagesSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {

            WorkBook m_book = new WorkBook();

            //Inserting image
            m_book.addPicture(1, 0, 3, 8, "..\\template\\neza.jpg");

            m_book.write(".\\Sample.xls");


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
                return "src\\Images\\WriteImagesSample.cs";
            }
        }
    }
}
