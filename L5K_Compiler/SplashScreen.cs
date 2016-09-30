using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace L5K_Compiler
{
    public partial class SplashScreen : Form
    {
        private double m_dblOpacityIncrement = .05;
        private double m_dblOpacityDecrement = .08;
        private const int TIMER_INTERVAL = 50;
        public SplashScreen()
        {
            InitializeComponent();
        }
        static SplashScreen ms_frmSplash = null;

        static private void ShowForm()
        {
            ms_frmSplash = new SplashScreen();
            Application.Run(ms_frmSplash);
        }

        static public void CloseForm()
        {
            System.Threading.Thread.Sleep(1500);
            if (ms_frmSplash != null)
            {
                ms_frmSplash.m_dblOpacityIncrement = -ms_frmSplash.m_dblOpacityDecrement;
            }
            ms_oThread = null;
            ms_frmSplash = null;
        }

        static Thread ms_oThread = null;

        static public void ShowSplashScreen()
        {
            // Make sure it is only launched once.
            if (ms_frmSplash != null)
                return;
            ms_oThread = new Thread(new ThreadStart(SplashScreen.ShowForm));
            ms_oThread.IsBackground = true;
            ms_oThread.SetApartmentState(ApartmentState.STA);
            ms_oThread.Start();
            while (ms_frmSplash == null || ms_frmSplash.IsHandleCreated == false)
            {
                System.Threading.Thread.Sleep(TIMER_INTERVAL);
            }
        }

        private string m_sStatus;

        static public void SetStatus(string newStatus)
        {
            if (ms_frmSplash == null)
                return;
            ms_frmSplash.m_sStatus = newStatus;
        }

        private double m_Progress;

        static public void SetProgress(double newProgress)
        {
            if (ms_frmSplash == null)
                return;
            ms_frmSplash.m_Progress = newProgress;
        }

        private void UpdateTimer_Tick(object sender, System.EventArgs e)
        {
            progBar.Value = (int)m_Progress;
            lblStatus.Text = m_sStatus;
            if (m_dblOpacityIncrement > 0)
            {
                if (this.Opacity < 100 )
                    this.Opacity += m_dblOpacityIncrement;
            }
            else
            {
                if (this.Opacity > 0)
                    this.Opacity += m_dblOpacityIncrement;
                else
                    this.Close();
            }
        }
    }
}
