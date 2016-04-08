using System;
using SmartXLS;

namespace SmartXLSExplorer
{

    public class DataValidationReadSample : BaseFileViewControl, ISampleControl
    {
        public DataValidationReadSample()
        {
            InitializeComponent();
            buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook book = new WorkBook();

            //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
            //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
            book.read("..\\template\\DVTemplate.xls");

            DataValidation validation = book.getValidation(0, 0);

            //Reading the Data Validation list
            string lists = validation.Formula1;

            lists = lists.Replace("\0", " ");
         
            RichTextForm textform = new RichTextForm();
            textform.richTextBox1.Text = lists;
            textform.ShowDialog();
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\DataValidation\\DataValidationReadSample.cs";
            }
        }
    }
}
