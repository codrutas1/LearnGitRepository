using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace SmartXLSExplorer
{
    public partial class SmartXLSExplorerControl : UserControl
    {
        public SmartXLSExplorerControl()
        {
            InitializeComponent();
        }

        private void navigationControl1_SelectionChanged(object sender, TreeViewEventArgs e)
        {
            tableLayoutPanel1.AutoSize = false;
            TreeNode node = e.Node;
            string nodeName = node.Name;
            string nodeText = node.Text;
            string typeName = "SmartXLSExplorer." + nodeName;
            if (nodeName.Contains("Sample"))
            {
                string codeText;
                ISampleControl sampleControl = sampleContainer1.LoadSampleControl(typeName);
                if (sampleControl != null)
                {
                    string codePath = sampleControl.CodePath;
                    if (codePath != null)
                    {
                        try
                        {
                            System.Diagnostics.Debug.Assert(codePath.Contains(nodeName));
                            string path = "..\\" + codePath;
                            System.IO.StreamReader reader = new System.IO.StreamReader(path);
                            using (reader)
                            {
                                codeText = reader.ReadToEnd();
                                codeText = codeText.Replace("<", "&lt;");
                                codeText = codeText.Replace(">", "&gt;");
                                codeText = codeText.Replace("\"", "&quot;");
                                codeText = codeText.Replace(" ", "&nbsp;");
                                codeText = codeText.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;");
                                codeText = codeText.Replace("\r\n", "<br />");
                                reader.Close();
                            }
                        }
                        catch (Exception exc)
                        {
                            codeText = "Exception parsing code text Exception=" + exc.ToString();
                        }
                    }
                    else
                        codeText = "Missing code file!";
                }
                else
                    codeText = "Missing sample control!";
                string htmlText = "<html>";
                htmlText += "<body style=\"background-color: #6666cc\">";
                htmlText += "<table border=\"0\" cellpadding=\"4\" cellspacing=\"6\" width=\"100%\" style=\"background-color: Window\">";
                htmlText += "<tr><td style=\"color: white; background-color: #6666cc; font-size: medium; font-weight: bold\">" + nodeText + " Sample - C#</td></tr>";
                htmlText += "<tr><td style=\"font-family: Courier New; font-size: x-small\">" + codeText + "</td></tr>";
                htmlText += "</table>";
                htmlText += "</body></html>";
                samplePanel.DocumentText = htmlText;
            }
            else
            {
                sampleContainer1.ClearSampleControl();
                string htmlText = "<html><body style=\"background-color: #6666cc\">";
                htmlText += "<table border=\"0\" cellpadding=\"4\" cellspacing=\"6\" width=\"100%\" style=\"background-color: Window\">";
                htmlText += "<tr><td style=\"color: white; background-color: #6666cc; font-size: medium; font-weight: bold\">" + nodeText + " Category</td></tr>";
                htmlText += "<tr><td style=\"font-family: Courier New; font-size: x-small\">" + "" + "</td></tr>";
                htmlText += "</table>";
                htmlText += "</body></html>";
                samplePanel.DocumentText = htmlText;
            }
            tableLayoutPanel1.AutoSize = true;
        }
    }
}
