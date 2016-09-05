using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Logging
{
    public class MessageList: IList<String>
    {
        IList<String> innerCollection = new List<String>();
        public MessageList()
        { 
        }

        public override String ToString()
        {
            String newString = string.Empty;

            foreach (String line in innerCollection)
            {
                newString = newString + line + Environment.NewLine;
            }

            return newString;
        }

        public String this[Int32 index]
        {
            get { return innerCollection[index]; }
            set { innerCollection[index] = value; }
        }

        public void Add(String item)
        {
            innerCollection.Add(item);
        }

        public void Clear()
        {
            innerCollection.Clear();
        }

        public Boolean Contains(String item)
        {
            return innerCollection.Contains(item);
        }

        public Boolean IsReadOnly
        {
            get { return innerCollection.IsReadOnly;}
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            return GetEnumerator();
        }

        System.Collections.Generic.IEnumerator<String> System.Collections.Generic.IEnumerable<String>.GetEnumerator()
        {
            return innerCollection.GetEnumerator();
        }

        public void CopyTo(String[] array, int arrayIndex)
        {
            innerCollection.CopyTo(array, arrayIndex);
        }

        public void RemoveAt(Int32 index)
        {
            innerCollection.RemoveAt(index);
        }

        public void Insert(Int32 index, String item)
        {
            innerCollection.Insert(index, item);
        }

        public Int32 IndexOf(String item)
        {
            return innerCollection.IndexOf(item);
        }

        public Int32 Count
        {
            get { return innerCollection.Count; }
        }

        public Boolean Remove(String item)
        {
            return innerCollection.Remove(item);
        }

        public IList<String> GetDeepCopy()
        {
            return Reusable.Collections.Generic.DeepCopy(innerCollection);
        }
    }
}
