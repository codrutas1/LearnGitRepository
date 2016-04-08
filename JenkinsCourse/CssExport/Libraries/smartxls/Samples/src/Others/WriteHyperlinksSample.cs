using System;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class WriteHyperlinksSample : BaseFileViewControl, ISampleControl
    {
        public WriteHyperlinksSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {

            try
            {
                WorkBook workBook = new WorkBook();

                //add a url link to F6
                workBook.addHyperlink(5, 5, 5, 5, "http://www.smartxls.com/", HyperLink.kURLAbs, "Hello,web url hyperlink!");

                //add a file link to F7
                workBook.addHyperlink(6, 5, 6, 5, "c:\\", HyperLink.kFileAbs, "file link");

                workBook.write(".\\Sample.xls");
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
                return "src\\Others\\WriteHyperlinksSample.cs";
            }
        }
    }
}
