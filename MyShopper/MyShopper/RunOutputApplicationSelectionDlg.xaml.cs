/*************************************
*  Created by [BOSS] Game Developers *
*     All Rights Reserved ©2017      *
*      [BOSS] Game Developers        *
*        15671 E Colorado Ave        *
*         Aurora, CO 80017           *
*    bossgamesdevteam@gmail.com      *
*         (720) 322-6242             *
*************************************/

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.ComponentModel;

namespace MyShopper
{
    /// <summary>
    /// Interaction logic for RunOutputApplicationSelectionDlg.xaml
    /// </summary>
    public partial class RunOutputApplicationSelectionDlg : Window
    {
        public bool m_bWord;
        public bool m_bExcel;
        public bool m_bNotepad;
        public bool m_bActive;

        public RunOutputApplicationSelectionDlg()
        {
            InitializeComponent();
            m_bExcel = m_bNotepad = m_bWord = false;

            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Closing += OnWindowClosing;

            m_bActive = true;

            IntPtr pIcon = Properties.Resources.Cart_Icon.ToBitmap().GetHbitmap();
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            this.Icon = pBitmap;

        }

        /// <summary>
        /// Close Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnWindowClosing(object sender, CancelEventArgs args)
        {
            m_bActive = false;
        }

        /// <summary>
        /// Retrieve the users selection for output application.
        /// If none are selected, MS Word is default
        /// </summary>
        private void getUserSelection()
        {
            foreach (Control control in spOpts.Children)
            {
                if (control.GetType() == typeof(RadioButton))
                {
                    if (((RadioButton)control).IsChecked == true)
                    {
                        string sVal = ((RadioButton)control).Content.ToString();
                        if (!string.IsNullOrEmpty(sVal))
                        {
                            if (sVal == "Notepad")
                                m_bNotepad = true;
                            else if (sVal == "MS Excel")
                                m_bExcel = true;
                            else m_bWord = true;
                        }
                        else m_bWord = true;
                    }
                }
            }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            getUserSelection();
            this.Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Escape:
                    this.Close();
                    break;
                case Key.Enter:
                    btnOk_Click(sender, e);
                    break;
            }
        }
    }
}
