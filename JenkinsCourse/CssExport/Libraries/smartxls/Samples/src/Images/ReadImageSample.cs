using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;



namespace SmartXLSExplorer
{
    public class ReadImageSample : BaseFileViewControl, ISampleControl
    {
        public ReadImageSample()
        {
            InitializeComponent();
            buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = new WorkBook();

            //open the workbook
            try
            {

                m_book.read("..\\template\\book.xls");
            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
            PictureShape picShape = m_book.getPictureShape(0);

            byte[] imagedata = picShape.PictureImageData;

            System.IO.FileStream fos = new System.IO.FileStream("pic.jpg", System.IO.FileMode.Create);
            fos.Write(imagedata, 0, imagedata.Length);
            fos.Close();

            System.Diagnostics.Process.Start("pic.jpg");
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Images\\ReadImageSample.cs";
            }
        }
    }
}
