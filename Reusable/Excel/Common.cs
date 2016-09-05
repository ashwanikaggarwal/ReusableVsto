using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ObjectModel = System.Collections.ObjectModel;

namespace Reusable.Excel
{
    public class UI
    {
        /// <summary>
        /// static method for validating entered value against combo box values
        /// </summary>
        /// <param name="control"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool ValidateComboBox(ComboBox control, string value)
        {
            bool ItemFound = false;
            foreach (object itm in control.Items)
            {
                string currentValue = itm.ToString();
                if (currentValue == value)
                {
                    ItemFound = true;
                }
            }
            return ItemFound;
        }
    }
}
