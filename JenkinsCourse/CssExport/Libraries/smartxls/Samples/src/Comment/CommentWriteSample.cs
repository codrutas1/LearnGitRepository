using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class CommentWriteSample : BaseFileViewControl, ISampleControl
    {
        public CommentWriteSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                //add a comment to B2
                workBook.addComment(1, 1, "comment text here!", "author name here!");

                workBook.write(".\\comment.xls");

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(".\\comment.xls");
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Comment\\CommentWriteSample.cs";
            }
        }
    }
}
