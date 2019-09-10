using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public class Reader
    {
        string _basePath;
        public Reader(string basePath)
        {
            _basePath = basePath;
        }
        public List<Paragraph> FindAllIfParagraphs(string filePath)
        {
            List<Paragraph> paragraphs = new List<Paragraph>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(" IF ")))
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    Console.WriteLine(paragraph.Count());
                    paragraphs.Add(paragraph);
                }
            }
            return paragraphs;
        }

        public void ReadLeapForm(string myFilePath)
        {
            // Use spire to get a collection of all mergefield names
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(myFilePath);

            var merge = document.MailMerge;
            FieldCollection fields = document.Fields;
            var mergeFields = from Spire.Doc.Fields.Field f in fields
                              where f.DocumentObjectType.ToString() == "MergeField"
                              select f;
            mergeFields = mergeFields.ToList();

            var mergeFieldNames = document.MailMerge.GetMergeFieldNames().ToList();

            foreach (var f in mergeFields)
            {
                Console.Write($"{f.Code.Substring(12, f.Code.Length - 12)}, {f.Text}\n");
            }
        }

        public void ReadSmokeballForm()
        {
            // Use spire to get a collection of fields in the smokeball document
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile($"{_basePath}SmokeballBankruptcyPetition.docx");

            var merge = document.MailMerge;
            FieldCollection fields = document.Fields;
            var mergeFields = from Spire.Doc.Fields.Field f in fields
                              select f;
            mergeFields = mergeFields.ToList();

            foreach (var f in mergeFields)
            {
                Console.Write($"{f.Code}, {f.Text}\n");
            }
        }
    }
}
