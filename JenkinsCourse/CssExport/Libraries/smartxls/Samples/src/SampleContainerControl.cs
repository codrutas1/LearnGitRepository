using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace SmartXLSExplorer
{

    public partial class SampleContainerControl : UserControl
    {
        private SampleManager _sampleManager;
        private ISampleControl _sampleControl;

        public SampleContainerControl()
        {
            InitializeComponent();
            _sampleManager = new SampleContainerControl.SampleManager();
        }

        internal ISampleControl LoadSampleControl(string typeName)
        {
            ClearSampleControl();
            Control sampleControl = _sampleManager.GetSampleControl(typeName);
            AddSampleControl(sampleControl);
            _sampleControl = (ISampleControl)sampleControl;
            return _sampleControl;
        }

        private void AddSampleControl(Control sampleControl)
        {
            if (sampleControl != null)
            {
                System.Diagnostics.Debug.Assert(Controls.Count == 0);
                sampleControl.Dock = DockStyle.Fill;
                Controls.Add(sampleControl);
            }
        }

        internal void ClearSampleControl()
        {
            if (Controls.Count > 0)
            {
                System.Diagnostics.Debug.Assert(Controls.Count == 1);
                Control sampleControl = Controls[0];
                sampleControl.Dock = DockStyle.None;
                sampleControl.SetBounds(0, 0, 0, 0);
                Controls.Clear();
            }
        }

        private class SampleManager
        {
            private List<UserControl> _sampleControls;

            internal SampleManager()
            {
                _sampleControls = new List<UserControl>();
            }

            internal UserControl FindControl(System.Type type)
            {
                for (int i = 0; i < _sampleControls.Count; i++)
                {
                    UserControl sample = _sampleControls[i];
                    if (sample.GetType() == type)
                        return sample;
                }
                return null;
            }

            internal UserControl GetSampleControl(string typeName)
            {
                UserControl sampleControl = null;
                if (typeName != null)
                {
                    System.Type type = System.Type.GetType(typeName);
                    sampleControl = FindControl(type);
                    if (sampleControl == null)
                    {
                        try
                        {
                            ConstructorInfo[] constructorInfos = type.GetConstructors();
                            System.Diagnostics.Debug.Assert(constructorInfos.Length == 1);
                            sampleControl = (UserControl)constructorInfos[0].Invoke(null);
                            if (sampleControl != null)
                                _sampleControls.Add(sampleControl);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                return sampleControl;
            }
        }
    }

    internal interface ISampleControl
    {
        string CodePath
        {
            get;
        }
    }
}
