using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Common
{
    public class Methods
    {
        /// <summary>
        /// Method used for Excel, to get column letter from column index
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Int32 GetColumnindex(char value)
        {
            string columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            Int32 columnIndex = 0;
            columnIndex = columns.IndexOf(value);
            return columnIndex + 1;
        }
    }
}
