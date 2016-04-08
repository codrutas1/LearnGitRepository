using System;
using System.Windows.Forms;

namespace SmartXLSExplorer
{
    public partial class SmartXLSExplorerApplication : Form
    {
        public SmartXLSExplorerApplication()
        {
            InitializeComponent();
        }
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SmartXLSExplorerApplication());
        }
    }
}