using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class DrawingObjectsSample : BaseFileViewControl, ISampleControl
    {
        public DrawingObjectsSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook workBook = new WorkBook();

            try
            {
                workBook.setText(0, 0, "aa");
                workBook.setText(1, 0, "bb");
                workBook.setText(2, 0, "cc");

                FormControlShape checkBoxShape = workBook.addFormControl(3.0, 1.0, 4.1, 2.1, FormControlShape.CheckBox);
                checkBoxShape.CellRange = "A1:A3";
                checkBoxShape.CellLink = "B2";
                checkBoxShape.Text = "checkbox1";

                FormControlShape comBoxShape1 = workBook.addFormControl(3.0, 3.0, 4.1, 4.1, FormControlShape.CombBox);
                comBoxShape1.CellRange = "A1:A3";
                comBoxShape1.CellLink = "B4";

                FormControlShape formControlShape2 = workBook.addFormControl(11.0, 2.0, 13.4, 5.5, FormControlShape.ListBox);
                formControlShape2.CellRange = "A1:A3";
                formControlShape2.CellLink = "B4";
             
                AutoShape rectangleShape = workBook.addAutoShape(1.0, 9.0, 3.0, 11.0, AutoShape.Rectangle);
                AutoShape textBoxShape = workBook.addAutoShape(3.0, 5.0, 5.0, 7.0, AutoShape.TextBox);
                textBoxShape.Text = "textBox ddd";
                AutoShape ovalShape = workBook.addAutoShape(3.5, 9.0, 5.0, 11.0, AutoShape.Oval);
                AutoShape lineShape = workBook.addAutoShape(6.0, 9.0, 8.0, 11.0, AutoShape.Line);

                workBook.write(".\\Sample.xls");
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
                return "src\\Others\\DrawingObjectsSample.cs";
            }
        }
    }
}
