/*************************************
*  Created by [BOSS] Game Developers *
*     All Rights Reserved ©2017      *
*      [BOSS] Game Developers        *
*     2172 S Trenton Way #5-206      *
*         Denver, CO 80231           *
*    bossgamesdevteam@gmail.com      *
*         (516) 302 - 3680           *
*************************************/


namespace MyShopper
{
    class ShoppingItem
    {
        public string sItemName { get; set; }
        public string Volume { get; set; }
        public bool bIsCustomsItem { get; set; }
        public bool bPounds { get; set; }
        public bool bOunches { get; set; }
        public int Quantity { get; set; }
        public double Mass { get; set; }

        #region Contructors
        public ShoppingItem() { }
        public ShoppingItem(string s)
        {
            sItemName = s;
        }
        public ShoppingItem(string s, int i)
        {
            sItemName = s;
            Quantity = i;
        }
        public ShoppingItem(string s, double d)
        {
            sItemName = s;
            Mass = d;
        }
        public ShoppingItem(string s, string vol)
        {
            sItemName = s;
            Volume = vol;
        }
        public ShoppingItem(string s, int i, string vol)
        {
            sItemName = s;
            Quantity = i;
            Volume = vol;
        }
        public ShoppingItem(string s, string vol, double d)
        {
            sItemName = s;
            Volume = vol;
            Mass = d;
        }
        public ShoppingItem(string s, int i, double d)
        {
            sItemName = s;
            Quantity = i;
            Mass = d;
        }
        public ShoppingItem(string s, int i, double d, string vol)
        {
            sItemName = s;
            Volume = vol;
            Mass = d;
            Volume = vol;
        }
        #endregion
        
    }
}
