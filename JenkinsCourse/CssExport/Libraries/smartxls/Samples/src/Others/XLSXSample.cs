using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class XLSXSample : BaseFileViewControl, ISampleControl
    {
        public XLSXSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = new WorkBook();

            try
            {
                m_book.readXLSX("..\\template\\format.xlsx");
                m_book.writeXLSX(".\\Sample.xlsx");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(ex.Message);
            }

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes) 
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            else
            {
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Others\\XLSXSample.cs";
            }
        }
    }
}
