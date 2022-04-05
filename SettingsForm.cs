using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Lync.Model;

namespace LyncFellow
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();

            Icon = Properties.Resources.LyncFellow;
            Text += string.Format(" (Version {0})", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);

            switch (Properties.Settings.Default.RedOnDndCallBusy)
            {
                case ContactAvailability.DoNotDisturb:
                    OnDoNotDisturb.Checked = true;
                    break;
                case ContactAvailability.Busy:
                    OnBusy.Checked = true;
                    break;
                default:
                    OnCallConference.Checked = true;
                    break;
            }
            DanceOnIncomingCall.Checked = Properties.Settings.Default.DanceOnIncomingCall;

            //change button text to match AvailableColor
            FreeColor.Text = Properties.Settings.Default.AvailableColor;
 
        }

        private void GlueckkanjaLabel_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.themusicpeopleinc.com");
        }

        private void MyOwnWebsiteLabel_Click(object sender, EventArgs e)
        {
            Process.Start("http://glueckkanja.com/lyncfellow");
        }

        private void CloseButtton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.DanceOnIncomingCall = DanceOnIncomingCall.Checked;
            if (OnDoNotDisturb.Checked)
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.DoNotDisturb;
            }
            else if (OnBusy.Checked)
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.Busy;
            }
            else
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.None;
            }
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(FreeColor, 0, FreeColor.Height);
            contextMenuStrip1.Text = Properties.Settings.Default.AvailableColor;
            FreeColor.Text = Properties.Settings.Default.AvailableColor;

            //contextMenuStrip1.Show(button1, 0, button1.Height);
            //MessageBox.Show("YEAH");
            Trace.WriteLine("YEAH Color: " + Properties.Settings.Default.AvailableColor);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void greenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.AvailableColor = "Green";
            FreeColor.Text = Properties.Settings.Default.AvailableColor;
        }
 
        private void cyanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.AvailableColor = "Cyan";
            FreeColor.Text = Properties.Settings.Default.AvailableColor;
        }

        private void lilacToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.AvailableColor = "Lilac";
            FreeColor.Text = Properties.Settings.Default.AvailableColor;
        }
 
        private void label5_Click(object sender, EventArgs e)
        {
            Process.Start("http://tmppro.com");
        }

        private void label6_Click(object sender, EventArgs e)
        {
            Process.Start("http://on-stage.com");
        }
    }
}
