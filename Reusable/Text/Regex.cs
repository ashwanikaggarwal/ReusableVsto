using System;
using System.IO;
using System.Windows.Forms;
using System.Configuration;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

[assembly: CLSCompliant(true)]
namespace Reusable.Text
{
    public static class StaticMethods
    {
        public static string[] SplitString(string source, string splitter)
        {
            string[] strings = {""};

            try
            {
                strings = Regex.Split(source, splitter);
            }
            catch
            {
                //empty catch, but will return valid array with no values
            }
            return strings;
        }

        public static string ReplaceString(string source, string pattern, string replacement)
        {
            string newString = source;

            try
            {
                newString = Regex.Replace(source, pattern, replacement);
            }
            catch
            {
                //empty catch, but will return original string
            }
            return newString;
        }
    }
}