using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class ImportDataTableSample : BaseGridViewControl, ISampleControl
    {
        public ImportDataTableSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            //Imports Data from the Template spreadsheet into the Grid.

            WorkBook m_book = new WorkBook();
            m_book.read("..\\template\\NorthwindDataTemplate.xls");

            //Read data from spreadsheet.
            DataTable customersTable = m_book.ExportDataTable();
 
            //Data can be bind to DataGrid
            this.dataGridView.DataSource = customersTable;
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\ManipulatingData\\ImportDataTableSample.cs";
            }
        }
    }
}
