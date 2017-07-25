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
    /// Interaction logic for RunAboutDlg.xaml
    /// </summary>
    public partial class RunAboutDlg : Window
    {
        public bool m_bActive { get; set; }     //Do we have an active window already?

        public RunAboutDlg()
        {
            InitializeComponent();
            this.Closing += OnWindowClosing;
            var pStartLoc = WindowStartupLocation.CenterScreen;
            this.WindowStartupLocation = pStartLoc;
            m_bActive = true;

            IntPtr pIcon = Properties.Resources.Cart_Icon.ToBitmap().GetHbitmap();
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            this.Icon = pBitmap;
        }

        /// <summary>
        /// Window closing Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnWindowClosing(object sender, CancelEventArgs args)
        {
            m_bActive = false;
        }

        /// <summary>
        /// Load Logo in to ImageBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgLogo_Loaded(object sender, RoutedEventArgs e)
        {
            Image_Loaded(sender, Properties.Resources.BOSSGames.GetHbitmap());
        }

        /// <summary>
        /// Load Logo in to ImageBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="pIcon"></param>
        private void Image_Loaded(Object sender, IntPtr pIcon)
        {
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            Image pImage = sender as Image;
            pImage.Source = pBitmap;
        }

        /// <summary>
        /// KeyDown Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                this.Close();
        }
    }
}
