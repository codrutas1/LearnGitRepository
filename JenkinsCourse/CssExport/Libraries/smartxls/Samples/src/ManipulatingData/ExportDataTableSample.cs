using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class ExportDataTableSample : BaseGridViewControl, ISampleControl
    {
        public ExportDataTableSample()
            : base()
        {
            InitializeComponent();
            //Load Data
            DataTable northwindDt;
            DataSet customersDataSet = new DataSet();
            customersDataSet.ReadXml("..\\template\\Customers.xml", XmlReadMode.ReadSchema);
            northwindDt = customersDataSet.Tables[0];
            //Data can be bind to DataGrid
            this.dataGridView.DataSource = northwindDt;
        }

        //Exports the DataTable to a spreadsheet.
        public override void buttonView_Click(object sender, System.EventArgs e)
        {

            WorkBook m_book = new WorkBook();

            //Export DataTable.
            if (this.dataGridView.DataSource != null)
            {
                m_book.ImportDataTable((DataTable)this.dataGridView.DataSource, true, 0, 0, -1, -1);
            }
            else
            {
                MessageBox.Show("There is no datatable to export, Please import a datatable first", "Error");
                return;
            }

            //Saving the workbook to disk.
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
                return "src\\ManipulatingData\\ExportDataTableSample.cs";
            }
        }
    }
}
