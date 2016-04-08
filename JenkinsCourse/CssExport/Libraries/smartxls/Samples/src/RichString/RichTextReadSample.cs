using System;
using SmartXLS;

namespace SmartXLSExplorer
{

    public class RichTextReadSample : BaseFileViewControl, ISampleControl
    {
        public RichTextReadSample()
        {
            InitializeComponent();
            buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {

            WorkBook book = new WorkBook();

            book.read("..\\template\\book.xls");

            string rft = book.getRichText(0, 0);

            RichTextForm textform = new RichTextForm();
            textform.richTextBox1.Rtf = rft;
            textform.ShowDialog();
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\RichString\\RichTextReadSample.cs";
            }
        }
    }
}
