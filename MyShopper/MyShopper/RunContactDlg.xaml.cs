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
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.ComponentModel;

namespace MyShopper
{
    /// <summary>
    /// Interaction logic for RunContactDlg.xaml
    /// </summary>
    public partial class RunContactDlg : Window
    {
        private double m_rBreak;                        //Scrollviewer wait time
        private double m_rScrollOffset;                 //Position in our Scrollviewer
        private System.Timers.Timer m_pTimer;           //Timer for our Scrollviewer
        public bool m_bActive { get; set; }             //Do we already have an active window?

        /// <summary>
        /// Initialize window
        /// </summary>
        public RunContactDlg()
        {
            InitializeComponent();
            this.Closing += OnWindowClosing;
            var pStartLoc = WindowStartupLocation.CenterScreen;
            this.WindowStartupLocation = pStartLoc;
            m_bActive = true;

            IntPtr pIcon = Properties.Resources.Cart_Icon.ToBitmap().GetHbitmap();
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            this.Icon = pBitmap;

            string sMessage = "Here at [BOSS] Games, we try our best to make the best games and apps for you!  Unfortunately, sometimes we fall a little short." +
                "  That's why we encourage everyone to send us your feedback!  If there's something you want or need that we missed - LET US KNOW!\n\n" +
                "Email us at the address below with your feedback or requests and we'll do our best to put that in to our next version!  Sometimes it can be quick...  Sometimes it might take a while...  Just know that we're working hard to make the perfect games and apps for you!\n\n" +
                "We're sure we forgot a couple items...  Which of them can we add to this app for you?\n" +
                "Having some difficulties?  Let us know!\n" +
                "Want your Grocery List to work with a different exporter?  Let us know!\n\n";

            
            tbMessage.Text = sMessage;
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
        /// Load Logo in to ImageBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgContact_Loaded(object sender, RoutedEventArgs e)
        {
            Image_Loaded(sender, Properties.Resources.BOSSGames.GetHbitmap());
        }

        /// <summary>
        /// Load Logo in to ImageBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Image_Loaded(Object sender, IntPtr pIcon)
        {
            BitmapSource pBitmap = Imaging.CreateBitmapSourceFromHBitmap(pIcon, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            Image pImage = sender as Image;
            pImage.Source = pBitmap;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                this.Close();
        }
    }
}
