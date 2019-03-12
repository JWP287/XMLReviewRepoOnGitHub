using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;



using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace XMLBasicConsoleApp
{
    public class OpenXMLWord
    {
        static string filepath = "OpenXMLForWord.docm";
        static string msg = "Hallo JWP 222222";

        public static void CreateWordDoc()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledDocument))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text(msg));
            }
        }

        public static void Refresh()
        {
            string fwoe = Path.GetFileNameWithoutExtension(filepath);
            Directory.Delete(fwoe, true);
            System.IO.Compression.ZipFile.ExtractToDirectory(filepath, fwoe);
        }
    }
}
