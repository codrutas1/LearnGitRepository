using System;
using SmartXLS;

namespace SmartXLSExplorer
{

    public class CommentReadSample : BaseFileViewControl, ISampleControl
    {
        public CommentReadSample()
        {
            InitializeComponent();
            buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook workBook = new WorkBook();

            //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
            //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
            workBook.read("..\\template\\book.xls");

            // get the first index from the current sheet
            CommentShape commentShape = workBook.getComment(1, 7);
            if (commentShape != null)
            {
                //string text = "Author:" + commentShape.Author + "\n text:" + commentShape.Text;
                string text = commentShape.RichText;

                RichTextForm textform = new RichTextForm();
                textform.richTextBox1.Rtf = text;
                textform.ShowDialog();
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Comment\\CommentReadSample.cs";
            }
        }
    }
}
