using System;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class ReadHyperlinksSample : BaseFileViewControl, ISampleControl
    {
        public ReadHyperlinksSample()
            : base()
        {
            InitializeComponent();
            base.buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                workBook.read("..\\template\\book.xls");

                // get the first index from the current sheet
                HyperLink hyperLink = workBook.getHyperlink(0);
                string txtHyperlink = hyperLink.LinkString;
                Console.WriteLine(txtHyperlink);
                MessageBox.Show(txtHyperlink, "Hyperlink text is:");

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
                return "src\\Others\\ReadHyperlinksSample.cs";
            }
        }
    }
}
