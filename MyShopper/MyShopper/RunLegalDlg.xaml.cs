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
using System.Timers;

namespace MyShopper
{
    /// <summary>
    /// Interaction logic for RunLegalDlg.xaml
    /// </summary>
    public partial class RunLegalDlg : Window
    {
        private double m_rBreak;                        //Scrollviewer wait time
        private double m_rScrollOffset;                 //Position in our Scrollviewer
        private System.Timers.Timer m_pTimer;           //Timer for our Scrollviewer
        public bool m_bActive { get; set; }             //Do we already have an active window?

        /// <summary>
        /// Initialize window
        /// </summary>
        public RunLegalDlg()
        {
            InitializeComponent();
            this.Closing += OnWindowClosing;
            var pStartLoc = WindowStartupLocation.CenterScreen;
            this.WindowStartupLocation = pStartLoc;
            m_bActive = true;

            IntPtr pIcon = Properties.Resources.Cart_Icon.ToBitmap().GetHbitmap();
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            this.Icon = pBitmap;

            string sNotice = "[BOSS] Games\nDenver, Colorado\nbossgamedevteam@gmail.com\n\n\nLegal Notice:\n\n" +
                             "[BOSS] Games does not explicitly own any of the products listed in this application.\n\n" +
                             "All products and any Intellectual Property or otherwise reserved rights remain with the creator, manufacturer, or provider of the product.\n\n" +
                             "[BOSS] Games does not hold any claim to any merchandising or other rights associated with products listed within this application.\n\n" +
                             "All products listed in this application are strictly listed for the convenience of our users.\n\n" +
                             "[BOSS] Games reserves all rights including but not limited to Intellectual Property and source code as it pertains to this application and all other applications and games developed by[BOSS] Games.  The copying, manipulating and/or unauthorized use of said IP is punishable under State and Federal law, and will be adressed according to the severity of the infringment.\n\n\n\n";

            tbLegalNotice.Text = sNotice;
            m_rScrollOffset = 0;
            m_rBreak = 0;

            m_pTimer = new System.Timers.Timer();
            m_pTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimerEvent);
            m_pTimer.Interval = 100;
            m_pTimer.Enabled = true;
        }

        /// <summary>
        /// Timer.Tick Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTimerEvent(object sender, ElapsedEventArgs e)
        {
            scrollViewer.Dispatcher.Invoke(new Action(() =>
            {
                if (m_rBreak < 20) m_rBreak++;
                else
                {
                    m_rScrollOffset += 1;
                    if (m_rScrollOffset > scrollViewer.ScrollableHeight)
                    {
                        m_rScrollOffset = 0;
                        m_rBreak = 0;
                    }
                    scrollViewer.ScrollToVerticalOffset(m_rScrollOffset);
                }
            }));
        }

        /// <summary>
        /// Window Closing Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnWindowClosing(object sender, CancelEventArgs args)
        {
            m_pTimer.Enabled = false;
            m_bActive = false;
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

        /// <summary>
        /// Load Logo in to ImageBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgLegal_Loaded(object sender, RoutedEventArgs e)
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
    }
}
