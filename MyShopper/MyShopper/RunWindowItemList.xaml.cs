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
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.ComponentModel;
using System.Windows.Forms.Integration;
using System.IO;

namespace MyShopper
{
    /// <summary>
    /// Interaction logic for RinWindowItemList.xaml
    /// </summary>
    public partial class RunWindowItemList : Window
    {
        List<int> lstComboContent;                  //A List of int's for Qty ComboBoxes
        List<double> lstComboMassContent;           //A List of doubles for Mass ComboBoxes
        List<ShoppingItem> m_lItems;                //List of user-selected items 
        public bool m_bActive { get; set; }         //Do we already have an active window?

        public enum QUANTITY_FORMAT
        {
            INT,
            DOUBLE,
            STRING
        };              //What unit format are we looking for?

        public enum OUTPUT_APPLICATION
        {
            MS_WORD,
            MS_EXCEL,
            NOTEPAD
        };           //What application do we want to send out list to?


        /// <summary>
        /// Initialize the window
        /// </summary>
        public RunWindowItemList()
        {
            InitializeComponent();
            InitalizeLists();
            SetComboBoxes();
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
        /// Set content for every ComboBox
        /// </summary>
        private void SetComboBoxes()
        {
            for (int i = 0; i < 21; i++)
            {
                lstComboContent.Add(i);                                                     //FILL LIST WITH NUMBERS 0-10
            }
            for (double i = 0.0; i < 20.1; i += .5)
            {
                lstComboMassContent.Add(i);                                                 //FILL LIST WITH INCRIMENTS OF .5(LBS)
            }


            // ASSIGN LIST CONTENT TO EVERY COMBO BOX
            // FRUIT
            foreach (StackPanel spVert in spFruit.Children)
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;

            // VEGETABLES
            foreach (StackPanel spVert in spVeggie.Children)
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;

            //  DAIRY
            foreach (StackPanel sp in spCheese.Children)                                    //CHEESE
                foreach (DockPanel dp in sp.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spMilk.Children)                                       //MILK
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spCondiments.Children)                                 //CONDIMENTS
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;                  

            //  FROZEN
            foreach (StackPanel spVert in spFrozen.Children)
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;

            //  SNACKS
            foreach (DockPanel dp in spCrackers.Children)                                  //CRACKERS
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spCookies.Children)                                   //COOKIES
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spChips.Children)                                      //CHIPS
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spSnackOther.Children)                                 //OTHER
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;

            //  PANTRY
            foreach (DockPanel dp in spCleaning.Children)                                   //CLEANING SUPPLIES
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spBathroom.Children)                                   //BATHROOM SUPPLIES
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spBaking.Children)                                     //BAKING SUPPLIES
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (StackPanel sp in spPantryOther.Children)                               //OTHER
                foreach (DockPanel dp in sp.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;

            //  BEVERAGES
            foreach (DockPanel dp in spCoffee.Children)                                    //COFFEE
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spBevOther.Children)                                  //OTHER
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spJuice.Children)                                     //JUICE
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (StackPanel spVert in spSoda.Children)                                  //SODA
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (StackPanel spVert in spWine.Children)                                  //WINE
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;

            //  PROTEIN
            int itr = 1;
            foreach (DockPanel dp in spBeef.Children)                                      //BEEF
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboMassContent;
            foreach (DockPanel dp in spTurkey.Children)                                    //TURKEY
            {
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                    {
                        if (itr <= 3)                                                       //LBS OR QTY?
                            ((ComboBox)control).ItemsSource = lstComboMassContent;
                        else ((ComboBox)control).ItemsSource = lstComboContent;
                    }
                itr++;
            }
            itr = 1;                                                                        //RESET ITERATOR
            foreach (DockPanel dp in spChicken.Children)                                    //CHICKEN
            {
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                    {
                        if (itr == 1 || itr == 2 || itr == 4 || itr == 5)                   //LBS OR QTY?
                            ((ComboBox)control).ItemsSource = lstComboMassContent;
                        else ((ComboBox)control).ItemsSource = lstComboContent;
                    }
                itr++;
            }
            itr = 1;                                                                       //RESET ITERATOR
            foreach (DockPanel dp in spFish.Children)                                      //FISH
            {
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                    {
                        if (itr == 3 || itr == 5)                                          //LBS OR QTY?
                            ((ComboBox)control).ItemsSource = lstComboMassContent;
                        else ((ComboBox)control).ItemsSource = lstComboContent;
                    }
                itr += 1;
            }
            itr = 1;
            foreach (DockPanel dp in spPork.Children)                                      //PORK
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboMassContent;

            //CARBS
            foreach (DockPanel dp in spPasta.Children)                                      //PASTA
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spRice.Children)                                       //RICE
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spBread.Children)                                      //BREAD
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;

            //MISC
            foreach (DockPanel dp in spSpice.Children)                                      //SPICES
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spHomeGoods.Children)                                  //HOME GOODS
                foreach (var control in dp.Children)
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;
            foreach (DockPanel dp in spUtils.Children)                                      //UTENSILS AND 
                foreach (var control in dp.Children)                                        //SUPPLIES
                    if (control.GetType() == typeof(ComboBox))
                        ((ComboBox)control).ItemsSource = lstComboContent;

            //OTHER
            foreach (StackPanel spVert in spOther.Children)
                foreach (DockPanel dp in spVert.Children)
                    foreach (var control in dp.Children)
                        if (control.GetType() == typeof(ComboBox))
                            ((ComboBox)control).ItemsSource = lstComboContent;
        }

        /// <summary>
        /// Initialize all Lists
        /// </summary>
        private void InitalizeLists()
        {
            lstComboContent = new List<int>();
            lstComboMassContent = new List<double>();
            m_lItems = new List<ShoppingItem>();
        }

        /// <summary>
        /// Destroy all Lists
        /// </summary>
        private void DestructLists()
        {
            lstComboContent.Clear();
            lstComboContent = null;
            lstComboMassContent.Clear();
            lstComboMassContent = null;
            m_lItems.Clear();
            m_lItems = null;
        }

        /// <summary>
        /// Create a List of items based on user selections
        /// </summary>
        private void GetItemList()
        {
            string sNoTag = "";
            ShoppingItem pItem;
            // FRUIT
            foreach (StackPanel spVert in spFruit.Children)
                addItemsFromPanel(spVert, sNoTag, QUANTITY_FORMAT.INT);

            // VEGETABLES
            foreach (StackPanel spVert in spVeggie.Children)
                addItemsFromPanel(spVert, sNoTag, QUANTITY_FORMAT.INT);


            //  DAIRY
            foreach (StackPanel sp in spCheese.Children)                                        //CHEESE
                addItemsFromPanel(sp, " Cheese", QUANTITY_FORMAT.INT);

            addItemsFromPanel(spMilk, " Milk", QUANTITY_FORMAT.INT);                            //MILK

            //  FROZEN
            foreach (StackPanel spVert in spFrozen.Children)
                addItemsFromPanel(spVert, sNoTag, QUANTITY_FORMAT.INT);

            //  SNACKS
            addItemsFromPanel(spCrackers, sNoTag, QUANTITY_FORMAT.INT);
            foreach (DockPanel dp in spChips.Children)
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();
                foreach (Control control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                              //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                          //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                        //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName == "Cooler Ranch" || sItemName == "Original")
                                    sItemName += " Doritos";
                                pItem.sItemName = sItemName;
                            }

                        }
                    }
                    else if (control.GetType() == typeof(TextBox))
                    {
                        if (((TextBox)control).IsEnabled == true)
                        {
                            if (!string.IsNullOrEmpty(((TextBox)control).Text))
                            {
                                string sVal = ((TextBox)control).Text;
                                if (((TextBox)control) == tbDoritosO)
                                    sVal += " Doritos";
                                pItem.sItemName = sVal;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))
                    {
                        if (((ComboBox)control).IsEnabled == true)
                        {
                            if (((ComboBox)control).SelectedValue != null)
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else pItem.Quantity = 0;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            addItemsFromPanel(spCookies, sNoTag, QUANTITY_FORMAT.INT);
            foreach (DockPanel dp in spSnackOther.Children)
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();
                foreach (Control control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))
                    {
                        if (((CheckBox)control).IsChecked == true)
                        {
                            if (((CheckBox)control).Content != null)
                            {
                                string sVal = ((CheckBox)control).Content.ToString();
                                if (sVal == "Almonds" || sVal == "Cashews" ||
                                    sVal == "Peanuts" || sVal == "Pistachios")
                                {
                                    pItem.bOunches = true;
                                }
                                pItem.sItemName = sVal;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))
                    {
                        if (((TextBox)control).IsEnabled == true)
                        {
                            if (((TextBox)control).Text != null &&
                                    ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))
                    {
                        if (((ComboBox)control).IsEnabled == true)
                        {
                            if (((ComboBox)control).SelectedValue != null)
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else pItem.Quantity = 0;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }

            //  PANTRY
            foreach (DockPanel dp in spCleaning.Children)                                       //CLEANING SUPPLIES
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();

                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName == "Device")
                                    sItemName = "Swiffer " + sItemName;
                                else if (sItemName == "Dry" || sItemName == "Wet")
                                    sItemName = sItemName + " Swiffer Pads";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                pItem.Quantity = 0;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }

            addItemsFromPanel(spBathroom, sNoTag, QUANTITY_FORMAT.INT);                         //BATHROOM & HYGENE
            addItemsFromPanel(spBaking, sNoTag, QUANTITY_FORMAT.INT);                           //BAKING GOODS

            foreach (DockPanel dp in spCondiments.Children)                                     //CONDIMENTS
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();

                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName == "Balsamic" || sItemName == "Bleu Cheese" ||
                                    sItemName == "Caesar" || sItemName == "Honey Mustard" ||
                                    sItemName == "Pear Gorgonzola" || sItemName == "Ranch" ||
                                    sItemName == "Raspberry Vinaigrette" ||
                                    sItemName == "Thai Peanut")
                                    sItemName = sItemName + " Dressing";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))                              //IF THE CHECK BOX DOESN'T HAVE CONTENT 
                    {                                                                           //IT MUST BE A TEXT BOX...
                        if (((TextBox)control).IsEnabled == true)                               //IS THE TEXT BOX ENABLED?
                        {
                            if (((TextBox)control).Text != null &&                              //DOES THE TEXT BOX HAVE ANY CONTENT?
                                    ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                pItem.Quantity = 0;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            foreach (StackPanel spVert in spPantryOther.Children)                                   //OTHER
            {
                foreach (DockPanel dp in spVert.Children)
                {
                    if (dp.Children.Count < 3) continue;
                    pItem = new ShoppingItem();
                    foreach (Control control in dp.Children)
                    {
                        if (control.GetType() == typeof(CheckBox))
                        {
                            if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                            {
                                if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                                {
                                    string sItemName = ((CheckBox)control).Content as string;
                                    if (sItemName == "Kidney" ||
                                        sItemName == "Garbonzo" ||
                                        sItemName == "Pinto")
                                        sItemName += " Beans";
                                    pItem.sItemName = sItemName;
                                }
                            }
                        }
                        else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                        {
                            if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                            {
                                if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                                {
                                    int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                    pItem.Quantity = qty;
                                }
                                else if (((ComboBox)control).SelectedValue == null)
                                {
                                    pItem.Quantity = 0;
                                }
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(pItem.sItemName))
                        m_lItems.Add(pItem);
                }
            }

            //  BEVERAGES
            foreach (DockPanel dp in spCoffee.Children)                                         //COFFEE
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();

                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName == "Dark Roast" ||
                                    sItemName == "Flavored")
                                    sItemName = sItemName + " Coffee";
                                else if (sItemName == "Regular Coffee" || sItemName == "Black Tea")
                                { /*LEAVE THIS AS-IS*/ }
                                else sItemName = sItemName + " Tea";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))                              //IF THE CHECK BOX DOESN'T HAVE CONTENT 
                    {                                                                           //IT MUST BE A TEXT BOX...
                        if (((TextBox)control).IsEnabled == true)                               //IS THE TEXT BOX ENABLED?
                        {
                            if (((TextBox)control).Text != null &&                              //DOES THE TEXT BOX HAVE ANY CONTENT?
                                    ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                pItem.Quantity = 0;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }

            foreach (StackPanel spVert in spBeer.Children)                                      //BEER AND LIQUOR
            {
                pItem = new ShoppingItem();
                pItem.bIsCustomsItem = true;
                foreach (DockPanel dp in spVert.Children)
                {
                    foreach (Control control in dp.Children)
                    {
                        if (control.GetType() == typeof(TextBox))
                        {
                            if (((TextBox)control).IsEnabled == true)
                            {
                                if (((TextBox)control).Text != null &&
                                        ((TextBox)control).Text != "")
                                {
                                    string sItemName = ((TextBox)control).Text as string;
                                    pItem.sItemName = sItemName;
                                }
                            }
                        }
                        else if (control.GetType() == typeof(RadioButton))
                        {
                            if (((RadioButton)control).IsChecked == true)
                            {
                                string vol = ((RadioButton)control).Content.ToString();
                                pItem.Volume = vol;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            foreach (StackPanel spVert in spSoda.Children)                                      //SODA
                addItemsFromPanel(spVert, sNoTag, QUANTITY_FORMAT.INT);
            addItemsFromPanel(spJuice, " Juice", QUANTITY_FORMAT.INT);                          //JUICE
            foreach (StackPanel spVert in spWine.Children)                                      //WINE
                addItemsFromPanel(spVert, sNoTag, QUANTITY_FORMAT.INT);
            foreach (DockPanel dp in spBevOther.Children)                                       //OTHER
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName == "Flat" ||
                                    sItemName == "Sparkling")
                                    sItemName = sItemName + " Water";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))                              //IF THE CHECK BOX DOESN'T HAVE CONTENT 
                    {                                                                           //IT MUST BE A TEXT BOX...
                        if (((TextBox)control).IsEnabled == true)                               //IS THE TEXT BOX ENABLED?
                        {
                            if (((TextBox)control).Text != null &&                              //DOES THE TEXT BOX HAVE ANY CONTENT?
                                    ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                pItem.Quantity = 0;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }

            //  PROTEIN
            foreach (DockPanel dp in spBeef.Children)                                           //BEEF
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                if (sItemName != "Ground Beef")
                                    sItemName = sItemName + " (Steak)";
                                pItem.sItemName = sItemName;
                                pItem.bPounds = true;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                decimal dcm = 0;
                                try { dcm = Convert.ToDecimal(((ComboBox)control).SelectedValue); }
                                catch { }
                                double mass = Convert.ToDouble(dcm);
                                pItem.Mass = mass;
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                pItem.Mass = 0.0;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            var pCtr = 1;

            foreach (DockPanel dp in spFish.Children)                                           //FISH
            {
                pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))                              //IF THE CHECK BOX DOESN'T HAVE CONTENT 
                    {                                                                           //IT MUST BE A TEXT BOX...
                        if (((TextBox)control).IsEnabled == true)                               //IS THE TEXT BOX ENABLED?
                        {
                            if (((TextBox)control).Text != null &&                              //DOES THE TEXT BOX HAVE ANY CONTENT?
                                    ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (pCtr == 3 || pCtr == 5)
                            {
                                if (((ComboBox)control).SelectedValue != null)                  //IS THERE A SELECTED VALUE?
                                {
                                    decimal dcm = 0;
                                    try { dcm = Convert.ToDecimal(((ComboBox)control).SelectedValue); }
                                    catch { }
                                    double mass = Convert.ToDouble(dcm);
                                    pItem.Mass = mass;
                                    pItem.bPounds = true;
                                }
                                else if (((ComboBox)control).SelectedValue == null)
                                {
                                    pItem.Mass = 0.0;
                                    pItem.bPounds = true;
                                }
                            }
                            else
                            {
                                if (((ComboBox)control).SelectedValue != null)                  //IS THERE A SELECTED VALUE?
                                {
                                    int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                    pItem.Quantity = qty;
                                }
                                else if (((ComboBox)control).SelectedValue == null)
                                {
                                    pItem.Quantity = 0;
                                }
                            }
                        }
                    }
                }
                pCtr += 1;
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }

            pCtr = 1;                                                                           //RESET COUNTER
            foreach (DockPanel dp in spChicken.Children)                                        //CHICKEN
            {
                pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                sItemName = sItemName + " (Chicken)";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                             //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (pCtr == 1 || pCtr == 2 ||
                                pCtr == 4 || pCtr == 5)
                            {
                                if (((ComboBox)control).SelectedValue != null)                  //IS THERE A SELECTED VALUE?
                                {
                                    decimal dcm = 0;
                                    try { dcm = Convert.ToDecimal(((ComboBox)control).SelectedValue); }
                                    catch { }
                                    double mass = Convert.ToDouble(dcm);
                                    pItem.Mass = mass;
                                    pItem.bPounds = true;
                                }
                                else if (((ComboBox)control).SelectedValue == null)
                                {
                                    pItem.Mass = 0.0;
                                    pItem.bPounds = true;
                                }
                            }
                            else
                            {
                                if (((ComboBox)control).SelectedValue != null)                  //IS THERE A SELECTED VALUE?
                                {
                                    int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                    pItem.Quantity = qty;
                                }
                                else if (((ComboBox)control).SelectedValue == null)
                                {
                                    pItem.Quantity = 0;
                                }
                            }
                        }
                    }
                }
                pCtr += 1;
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            pCtr = 1;                                                                           //RESET COUNTER
            foreach (DockPanel dp in spTurkey.Children)                                         //TURKEY
            {
                pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                                  //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                              //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                            //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                sItemName = sItemName + " (Turkey)";
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    if (control.GetType() == typeof(ComboBox))                                  //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                              //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                      //IS THERE A SELECTED VALUE?
                            {
                                decimal dcm = 0;
                                try { dcm = Convert.ToDecimal(((ComboBox)control).SelectedValue); }
                                catch { }
                                if (pCtr <= 3)                                                  //CHECK TO SEE IF THE SELECTED 
                                {
                                    double mass = Convert.ToDouble(dcm);
                                    pItem.Mass = mass;
                                    pItem.bPounds = true;
                                }
                                else
                                {
                                    int qty = Convert.ToInt32(dcm);
                                    pItem.Quantity = qty;
                                }
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                if (pCtr <= 3)
                                {
                                    pItem.Mass = 0.0;
                                    pItem.bPounds = true;
                                }

                                else
                                    pItem.Quantity = 0;
                            }
                        }
                    }
                }
                pCtr += 1;
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            pCtr = 1;                                                                           //RESET COUNTER

            addItemsFromPanel(spPork, " (Pork)", QUANTITY_FORMAT.DOUBLE);

            //CARBS
            addItemsFromPanel(spPasta, " Pasta", QUANTITY_FORMAT.INT);                          //PASTA
            addItemsFromPanel(spRice, sNoTag, QUANTITY_FORMAT.INT);                             //RICE
            addItemsFromPanel(spBread, sNoTag, QUANTITY_FORMAT.INT);                            //BREAD

            //MISC
            addItemsFromPanel(spSpice, sNoTag, QUANTITY_FORMAT.INT);                            //SPICES
            foreach (DockPanel dp in spHomeGoods.Children)                                      //HOME GOODS
            {
                if (dp.Children.Count < 3) continue;
                pItem = new ShoppingItem();
                foreach (Control control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))
                    {
                        if (((CheckBox)control).IsChecked == true)
                        {
                            if (((CheckBox)control).Content != null)
                            {
                                string sVal = ((CheckBox)control).Content.ToString();
                                if (sVal == "Refill Packs")
                                {
                                    sVal = "Air Freshener " + sVal;
                                }
                                else if (sVal == "AA" || sVal == "AAA" || sVal == "C" || sVal == "D" ||
                                    sVal == "9V" || sVal == "123A" || sVal == "Watch")
                                {
                                    sVal += " Batteries";
                                }
                                else if (sVal == "Large" || sVal == "Small")
                                {
                                    sVal += " Ziploc Bags";
                                }
                                pItem.sItemName = sVal;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))
                    {
                        if (((TextBox)control).IsEnabled == true)
                        {
                            string sVal = ((TextBox)control).Text;
                            if (((TextBox)control) == tbBattO)
                            {
                                sVal = tbBattO.Text;
                                sVal += " Batteries";
                            }
                            pItem.sItemName = sVal;
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))
                    {
                        if (((ComboBox)control).IsEnabled == true)
                        {
                            if (((ComboBox)control).SelectedValue != null)
                            {
                                int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                pItem.Quantity = qty;
                            }
                            else pItem.Quantity = 0;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
            addItemsFromPanel(spUtils, sNoTag, QUANTITY_FORMAT.INT);                            //UTENSILS AND SUPPLIES
            addItemsFromPanel(spFirstAid, sNoTag, QUANTITY_FORMAT.INT);                         //FIRST AID ITEMS

            //OTHER GROUP BOX
            foreach (StackPanel spVert in spOther.Children)
            {
                foreach (DockPanel dp in spVert.Children)
                {
                    pItem = new ShoppingItem();
                    foreach (Control control in dp.Children)
                    {
                        if (control.GetType() == typeof(TextBox))
                        {
                            if (((TextBox)control).IsEnabled == true)
                            {
                                string sVal = ((TextBox)control).Text;
                                pItem.sItemName = sVal;
                            }
                        }
                        else if (control.GetType() == typeof(ComboBox))
                        {
                            if (((ComboBox)control).IsEnabled == true)
                            {
                                if (((ComboBox)control).SelectedValue != null)
                                {
                                    int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                    pItem.Quantity = qty;
                                }
                                else pItem.Quantity = 0;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(pItem.sItemName))
                        m_lItems.Add(pItem);
                }
            }
        }

        /// <summary>
        /// Gets our selected items from StackPanels and places them in our List<>
        /// </summary>
        /// <param name="sp">The StackPanel</param>
        /// <param name="sCustomTag">A custom tag, if the items need one (ex. Bread)</param>
        /// <param name="qFmt">The quantity format we're looking for</param>
        private void addItemsFromPanel(StackPanel sp, string sCustomTag, QUANTITY_FORMAT qFmt)
        {
            foreach (DockPanel dp in sp.Children)
            {
                if (dp.Children.Count < 3) continue;
                var pItem = new ShoppingItem();
                foreach (var control in dp.Children)
                {
                    if (control.GetType() == typeof(CheckBox))                              //ARE WE ON A CHECKBOX?
                    {
                        if (((CheckBox)control).IsChecked == true)                          //IS IT CHECKED?
                        {
                            if (((CheckBox)control).Content != null)                        //DOES IT HAVE CONTENT?
                            {
                                string sItemName = ((CheckBox)control).Content as string;
                                pItem.sItemName = sItemName + sCustomTag;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(TextBox))                          //IF THE CHECK BOX DOESN'T HAVE CONTENT 
                    {                                                                       //IT MUST BE A TEXT BOX...
                        if (((TextBox)control).IsEnabled == true)                           //IS THE TEXT BOX ENABLED?
                        {
                            if (((TextBox)control).Text != null &&                          //DOES THE TEXT BOX HAVE ANY CONTENT?
                                ((TextBox)control).Text != "")
                            {
                                string sItemName = ((TextBox)control).Text as string;
                                pItem.sItemName = sItemName;
                            }
                        }
                    }
                    else if (control.GetType() == typeof(ComboBox))                         //ARE WE ON A COMBO BOX?
                    {
                        if (((ComboBox)control).IsEnabled == true)                          //IS IT ENABLED?
                        {
                            if (((ComboBox)control).SelectedValue != null)                  //IS THERE A SELECTED VALUE?
                            {
                                switch (qFmt)
                                {
                                    case QUANTITY_FORMAT.INT:
                                        int qty = Convert.ToInt32(((ComboBox)control).SelectedValue);
                                        pItem.Quantity = qty;
                                        break;
                                    case QUANTITY_FORMAT.DOUBLE:
                                        decimal dcm = 0;
                                        try { dcm = Convert.ToDecimal(((ComboBox)control).SelectedValue); }
                                        catch { }
                                        double mass = Convert.ToDouble(dcm);
                                        pItem.Mass = mass;
                                        pItem.bPounds = true;
                                        break;
                                }
                            }
                            else if (((ComboBox)control).SelectedValue == null)
                            {
                                switch (qFmt)
                                {
                                    case QUANTITY_FORMAT.INT:
                                        pItem.Quantity = 0;
                                        break;
                                    case QUANTITY_FORMAT.DOUBLE:
                                        pItem.Mass = 0.0;
                                        pItem.bPounds = true;
                                        break;
                                }
                            }
                        }
                    }
                    else if (control.GetType() == typeof(RadioButton))                      //IS IT A RADIO BUTTON?
                    {
                        if (((RadioButton)control).IsEnabled == true)                       //IS IT ENABLED?
                        {
                            if (((RadioButton)control).IsChecked == true)                   //IS IT SELECTED?
                            {
                                string volume = ((RadioButton)control).Content.ToString();
                                pItem.Volume = volume;
                                pItem.bIsCustomsItem = true;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(pItem.sItemName))
                    m_lItems.Add(pItem);
            }
        }

        /// <summary>
        /// Create our Word Document, Excel Document or text file and export our List to it
        /// </summary>
        /// <param name="lstItems">Represents our full List of ShoppingItem's including name and Qty/Mass/Volume</param>
        /// This Master List should be created prior to this being called
        private void CreateDocument(List<ShoppingItem> lstItems)
        {
            RunOutputApplicationSelectionDlg pOpts = null;
            bool bWord, bExcel, bNotepad;
            OUTPUT_APPLICATION pApp;
            if (lstItems.Count < 1)
                MessageBox.Show("It's looks like you didn't select any items.  Either you're good to go - or we need to try this again...", "Oops!!");
            else
            {
                if (pOpts != null && pOpts.m_bActive) pOpts.Activate();
                else
                {
                    pOpts = new RunOutputApplicationSelectionDlg();
                    ElementHost.EnableModelessKeyboardInterop(pOpts);
                    pOpts.ShowDialog();
                }
                bWord = pOpts.m_bWord;
                bExcel = pOpts.m_bExcel;
                bNotepad = pOpts.m_bNotepad;

                if (bExcel) pApp = OUTPUT_APPLICATION.MS_EXCEL;
                else if (bNotepad) pApp = OUTPUT_APPLICATION.NOTEPAD;
                else pApp = OUTPUT_APPLICATION.MS_WORD;

                int itr = 0;
                string[] sList = new string[lstItems.Count];

                foreach (ShoppingItem sItem in lstItems)                                        //LET'S LOOK AT ALL OF OUR ITEMS
                {
                    string sUnit;
                    if (sItem.bOunches) sUnit = "oz.(s)";
                    else if (sItem.bPounds) sUnit = "lb.(s)";
                    else sUnit = "";
                    string sQty = sItem.Quantity.ToString();                                   //LET'S FIND OUR QUANTITY
                    if (sQty == null || sQty == "0")
                        sQty = sItem.Mass.ToString();
                    if (sQty == null || sQty == "0")
                    {
                        if (sItem.Volume != null)
                            sQty = sItem.Volume;
                        else sQty = "(No Measure Selected)";
                    }

                    sList[itr] = sItem.sItemName + "\t--\t" + sQty + " " + sUnit + "\n";        //MAKE A STRING OF THE ITEM
                    itr++;                                                                      //GO TO THE NEXT ONE
                }
                itr = 0;

                switch (pApp)
                {
                    case OUTPUT_APPLICATION.MS_WORD:
                        var pWordApp = new Word.Application();                                          //CREATE OUR WORD APP
                        pWordApp.Visible = true;                                                        //MAKE SURE WE CAN SEE IT
                        Word.Document pDoc = pWordApp.Documents.Add();                                  //ADD A DOCUMENT TO IT SO WE HAVE SOMETHING TO WRITE ON
                        pDoc.PageSetup.TextColumns.SetCount(2);

                        try
                        {
                            for (; itr < sList.Count(); itr++)
                                pWordApp.Selection.InsertAfter(sList[itr]);                             //SEND THE ITEM TO THE WORD DOC
                        }
                        catch
                        {
                            pWordApp.ActiveDocument.Close();
                            MessageBox.Show("Your list is empty or could not be created!", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                            //NEEDS TO ALSO CLOSE MS WORD.  DOESN'T AND I CAN'T FIND A WAY SO FAR...
                            pWordApp.ActiveWindow.Close();
                        }

                        break;

                    case OUTPUT_APPLICATION.MS_EXCEL:
                        var pExcelApp = new Excel.Application();
                        pExcelApp.Visible = true;
                        Excel.Workbook pWorkbook = pExcelApp.Workbooks.Add();
                        Excel.Worksheet pWorksheet = (Excel.Worksheet)pWorkbook.Worksheets[1];
                        pWorksheet.Columns.ColumnWidth = 30;

                        int iRow = 1;
                        string sCell1, sCell2;
                        sCell1 = "A1";
                        sCell2 = "B1";
                        var pRange = pWorksheet.get_Range(sCell1, sCell1);
                        pRange.Font.Bold = true;
                        pRange.Font.Underline = true;

                        Object[] sItem = new Object[1];
                        sItem[0] = "Items";

                        pRange.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, pRange, sItem);

                        sItem[0] = "Quantity";
                        pRange = pWorksheet.get_Range(sCell2, sCell2);
                        pRange.Font.Bold = true;
                        pRange.Font.Underline = true;
                        pRange.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, pRange, sItem);

                        foreach (ShoppingItem pItem in lstItems)
                        {
                            iRow++;
                            sCell1 = "A" + iRow.ToString();
                            sCell2 = "B" + iRow.ToString();

                            pRange = pWorksheet.get_Range(sCell1, sCell1);
                            sItem[0] = pItem.sItemName;
                            pRange.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, pRange, sItem);


                            pRange = pWorksheet.get_Range(sCell2, sCell2);
                            string sUnit;
                            if (pItem.bOunches) sUnit = "oz.(s)";
                            else if (pItem.bPounds) sUnit = "lb.(s)";
                            else sUnit = "";
                            string sQty = pItem.Quantity.ToString();                                   //LET'S FIND OUR QUANTITY
                            if (sQty == null || sQty == "0")
                                sQty = pItem.Mass.ToString();
                            if (sQty == null || sQty == "0")
                            {
                                if (pItem.Volume != null)
                                    sQty = pItem.Volume;
                                else sQty = "(No Measure Selected)";
                            }

                            sQty = sQty + " " + sUnit;
                            sItem[0] = sQty;

                            pRange.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, pRange, sItem);
                        }

                        break;

                    case OUTPUT_APPLICATION.NOTEPAD:
                        string sUserName = Environment.UserName;
                        string sFilepath = @"C:\Users\" + sUserName + @"\Desktop\My Grocery List.txt";
                        using (StreamWriter pFile = new StreamWriter(sFilepath))
                        {
                            pFile.WriteLine("My Grocery List:");
                            pFile.WriteLine("");
                            foreach (string sLine in sList)
                            {
                                pFile.WriteLine(sLine);
                            }
                        }

                        MessageBox.Show("Your list has been saved to your desktop!");
                        break;
                }

            }


        }

        /// <summary>
        /// Create List Button Click Event Handler
        /// </summary>
        private void btnCreateList_Click(object sender, RoutedEventArgs e)
        {
            GetItemList();                                  //COMPILE MASTER LIST OF ITEMS AND QUANTITIES/VOLUMES
            CreateDocument(m_lItems);                       //SEND THE LIST TO A WORD DOCUMENT
            DestructLists();                                //DISPOSE OF ALL OF OUR LISTS
            this.Close();                                   //CLOSE THE WINDOW - WE'RE DONE HERE.

        }

        /// <summary>
        /// KeyDown Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                btnCreateList_Click(sender, e);
            if (e.Key == Key.Escape)
            {
                DestructLists();
                this.Close();
            }

        }

    }
}
