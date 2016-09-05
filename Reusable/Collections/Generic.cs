using System;
using System.Collections;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using ObjectModel = System.Collections.ObjectModel;

namespace Reusable.Collections
{
    class Generic
    {
        public static T DeepCopy<T>(T obj)
        {
            object result = null;

            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                result = (T)formatter.Deserialize(ms);
                ms.Close();
            }

            return (T)result;
        }

        /// <summary>
        /// returns collection of uniqe strings from comma-delimited string
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ObjectModel.Collection<string> UniqueCollectionFromCommaDelimitedString(string values)
        {
            string[] valueArray = values.Split(',');

            ObjectModel.Collection<string> filteredValues = new ObjectModel.Collection<string>();
            foreach (string item in valueArray)
            {
                if (!filteredValues.Contains(item.ToLower()))
                {
                    filteredValues.Add(item.ToLower().Trim());
                }
            }
            return filteredValues;
        }

        /// <summary>
        /// returns Hashtable of uniqe strings from comma-delimited string
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public static Hashtable UniqueHashtableFromCommaDelimitedString(string values)
        {
            string[] valueArray = values.Split(',');

            Hashtable filteredValues = new Hashtable(StringComparer.OrdinalIgnoreCase);
            foreach (string item in valueArray)
            {
                if (!filteredValues.Contains(item))
                {
                    filteredValues.Add(item, item);
                }
            }
            return filteredValues;
        }
    }
}
