using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace XMLBasicConsoleApp
{
    public class Sample1
    {
        private const String filename = "..\\..\\items.xml";

        public static void Run()
        {
            XmlTextReader reader = null;

            try
            {
                // Load the reader with the data file and ignore   
                // all white space nodes.  
                reader = new XmlTextReader(filename);
                reader.WhitespaceHandling = WhitespaceHandling.None;

                // Parse the file and display each of the nodes.  
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            Console.Write("<{0}>", reader.Name);
                            break;
                        case XmlNodeType.Text:
                            Console.Write(reader.Value);
                            break;
                        case XmlNodeType.CDATA:
                            Console.Write("<![CDATA[{0}]]>", reader.Value);
                            break;
                        case XmlNodeType.ProcessingInstruction:
                            Console.Write("<?{0} {1}?>", reader.Name, reader.Value);
                            break;
                        case XmlNodeType.Comment:
                            Console.Write("<!--{0}-->", reader.Value);
                            break;
                        case XmlNodeType.XmlDeclaration:
                            Console.Write("<?xml version='1.0'?>");
                            break;
                        case XmlNodeType.Document:
                            break;
                        case XmlNodeType.DocumentType:
                            Console.Write("<!DOCTYPE {0} [{1}]", reader.Name, reader.Value);
                            break;
                        case XmlNodeType.EntityReference:
                            Console.Write(reader.Name);
                            break;
                        case XmlNodeType.EndElement:
                            Console.Write("</{0}>", reader.Name);
                            break;
                    }
                    Console.WriteLine();
                }
            }

            finally
            {
                if (reader != null)
                    reader.Close();
            }
        }
        public static void Run2()
        {
            // Create the XmlDocument.  
            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" +
                        "<title>Pride And Prejudice</title>" +
                        "</book>");

            // Save the document to a file.  
            doc.Save("data.xml");
        }

        public static void Run3()
        {
            // Create the XmlDocument.  
            XmlDocument doc = new XmlDocument();
            doc.Load("data.xml");

            // Save the document to a file.  
            doc.Save("dataClone.xml");
        }

        public static void Run4()
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5' misc='sale item'>" +
                          "<title>The Handmaid's Tale</title>" +
                          "<price>14.95</price>" +
                          "</book>");

            // Move to an element.  
            XmlElement myElement = doc.DocumentElement;

            // Create an attribute collection from the element.  
            XmlAttributeCollection attrColl = myElement.Attributes;

            // Show the collection by iterating over it.  
            Console.WriteLine("Display all the attributes in the collection...");
            for (int i = 0; i < attrColl.Count; i++)
            {
                Console.Write("{0} = ", attrColl[i].Name);
                Console.Write("{0}", attrColl[i].Value);
                Console.WriteLine();
            }

            // Retrieve a single attribute from the collection; specifically, the  
            // attribute with the name "misc".  
            XmlAttribute attr = attrColl["misc"];

            // Retrieve the value from that attribute.  
            String miscValue = attr.InnerXml;

            Console.WriteLine("Display the attribute information.");
            Console.WriteLine(miscValue);

        }


        public static void RunXSL()
        {
            // Open books.xml as an XPathDocument.
            XPathDocument doc = new XPathDocument("books.xml");

            // Create a writer for writing the transformed file.
            XmlWriter writer = XmlWriter.Create("books.html");

            // Create and load the transform with script execution enabled.
            XslCompiledTransform transform = new XslCompiledTransform();
            XsltSettings settings = new XsltSettings();
            settings.EnableScript = true;
            transform.Load("transform.xsl", settings, null);

            // Execute the transformation.
            transform.Transform(doc, writer);
        }

    } // End class}
} 