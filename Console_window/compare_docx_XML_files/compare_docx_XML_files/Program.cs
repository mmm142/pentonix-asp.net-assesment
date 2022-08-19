using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
namespace compare_docx_XML_files
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("start");
            Console.ReadLine();
            string rootPath = @"D:\Project\Console_window";
            string xmlFile = rootPath + @"\XMLFile1.xml";
            string documentFile = rootPath + @"\DocFile1.doc";
            string outputDoc = rootPath + @"\MyGeneratedDocument.doc";

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentFile, true))
            {
                //get the main part of the document which contains CustomXMLParts
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                //delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable
                mainPart.DeleteParts<CustomXmlPart>(mainPart.CustomXmlParts);

                //add new CustomXMLPart with data from new XML file
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (FileStream stream = new FileStream(xmlFile, FileMode.Open))
                {
                    myXmlPart.FeedData(stream);
                }
            }

        }
    }
}
