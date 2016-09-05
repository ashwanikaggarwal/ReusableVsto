using System;
using System.Xml;
using XML = System.Xml;

namespace Reusable.XmlTools
{
    public class Xml
    {
        public static void ValidationEventHandler(object sender, XML.Schema.ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XML.Schema.XmlSeverityType.Error:
                    Console.WriteLine("Error: {0}", e.Message);
                    break;
                case XML.Schema.XmlSeverityType.Warning:
                    Console.WriteLine("Warning {0}", e.Message);
                    break;
            }
        }

        public static String EncodeAs(String currentXml, string encodeAsString)
        {
            System.Xml.XmlDocument oldDoc = new System.Xml.XmlDocument();
            oldDoc.LoadXml(currentXml);
            String strippedXml = oldDoc.DocumentElement.OuterXml;
            System.Xml.XmlDocument newDoc = new System.Xml.XmlDocument();
            System.Xml.XmlDeclaration xmldecl = newDoc.CreateXmlDeclaration("1.0", null, null);
            xmldecl.Encoding = encodeAsString;
            xmldecl.Standalone = "yes";
            newDoc.LoadXml(strippedXml);
            var root = newDoc.DocumentElement;
            newDoc.InsertBefore(xmldecl, root);
            return newDoc.InnerXml.ToString();
        }
    }
}
