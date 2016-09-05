using System;
using System.IO;

namespace Reusable.IO
{
    public class StreamTools
    {
        public static String ReadStream(string path)
        {
            String dataXml = string.Empty;
            using (StreamReader sr = new StreamReader(path))
            {
                dataXml = sr.ReadToEnd();
            }
            return dataXml;
        }
    }
}
