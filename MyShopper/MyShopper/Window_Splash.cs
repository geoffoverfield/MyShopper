/*************************************
*  Created by [BOSS] Game Developers *
*     All Rights Reserved ©2017      *
*      [BOSS] Game Developers        *
*     2172 S Trenton Way #5-206      *
*         Denver, CO 80231           *
*    bossgamesdevteam@gmail.com      *
*         (516) 302 - 3680           *
*************************************/

using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;


namespace MyShopper
{
    public partial class Window_Splash : Form
    {

        RunWindowItemList wWindow;
        RunAboutDlg wAbout;
        RunLegalDlg wLegal;
        RunContactDlg wContact;

        public Window_Splash()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        /// <summary>
        /// Run Main Program
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (wWindow != null && wWindow.m_bActive) wWindow.Activate();
            else
            {
                wWindow = new RunWindowItemList();
                ElementHost.EnableModelessKeyboardInterop(wWindow);
                wWindow.Show();
            }

        }

        /// <summary>
        /// Show our About window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAbout_Click(object sender, EventArgs e)
        {
            if (wAbout != null && wAbout.m_bActive) wAbout.Activate();
            else
            {
                wAbout = new RunAboutDlg();
                ElementHost.EnableModelessKeyboardInterop(wAbout);
                wAbout.Show();
            }
        }

        /// <summary>
        /// Show our legal notice
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLegal_Click(object sender, EventArgs e)
        {
            if (wLegal != null && wLegal.m_bActive) wLegal.Activate();
            else
            {
                wLegal = new RunLegalDlg();
                ElementHost.EnableModelessKeyboardInterop(wLegal);
                wLegal.Show();
            }
        }

        /// <summary>
        /// Show our Contact window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnContact_Click(object sender, EventArgs e)
        {
            if (wContact != null && wContact.m_bActive) wContact.Activate();
            else
            {
                wContact = new RunContactDlg();
                ElementHost.EnableModelessKeyboardInterop(wContact);
                wContact.Show();
            }
        }
    }
}
