using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public class ExperimentXML
    {
        public string WriteToWordDoc(string filePath, string txt)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Body body = doc.MainDocumentPart.Document.Body;

                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(txt));

                RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));
                run.AppendChild(new Text(txt));
            }

            return "Completed paragraph generation";
        }

        public string CreateWordDoc(string filePath, string text)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create document structure and add text
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // add text
                run.AppendChild(new Text(text));
            }

            return "New File Created";
        }
    }
}
