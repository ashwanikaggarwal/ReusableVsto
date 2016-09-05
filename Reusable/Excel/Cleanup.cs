using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Excel
{
    public class Com
    {
        /// <summary>
        /// sets COM object count to zero and sets to null for GC
        /// </summary>
        /// <param name="o"></param>
        public static void GarbageCleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        ///     Decrements COM object count and sets to null for GC
        /// </summary>
        /// <param name="o"></param>
        public static void ReleaseAndNull(object comObject)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
            }
            catch
            {
            }
            finally
            {
                comObject = null;
            }
        }

        /// <summary>
        /// sets COM object count to zero and sets to null for GC
        /// </summary>
        /// <param name="o"></param>
        public static void FinalReleaseAndNull(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }
    }
}
