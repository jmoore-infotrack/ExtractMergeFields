using DocumentFormat.OpenXml;
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
    public class MyMethodContainer
    {
        string _basePath;
        public MyMethodContainer(string basePath)
        {
            _basePath = basePath;
        }
        public void SaveZipFile(string openPath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(openPath, true))
            {
                string savePath = $"{_basePath}Petition.zip";
                doc.SaveAs(savePath);
            }
        }

        public void ExchangeTextValue(string filePath, string incorrectValue = "Harold", string correctValue = "Harry")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (Run run in doc.MainDocumentPart.Document.Descendants<Run>())
                {
                    foreach (Text txtFromRun in run.Descendants<Text>().Where(a => a.Text == incorrectValue))
                    {
                        txtFromRun.Text = correctValue;
                    }
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void MailMerge(string filePath, string incorrectValue = "Harold", string correctValue = "Harry")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                const string FieldDelimeter = " MERGEFIELD ";

                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                {
                    if (field.Text.Contains(FieldDelimeter.ToString()))
                    {
                        // var fieldNameStart = field.Text.LastIndexOf(FieldDelimeter, System.StringComparison.Ordinal);
                        // var fieldName = field.Text.Substring(fieldNameStart + FieldDelimeter.Length).Trim();

                        foreach (Run run in doc.MainDocumentPart.Document.Descendants<Run>())
                        {
                            foreach (Text txtFromRun in run.Descendants<Text>().Where(a => a.Text == incorrectValue))
                            {
                                txtFromRun.Text = correctValue;
                            }
                            //foreach (Text txtFromRun in run.Descendants<Text>())
                            //{
                            //    Console.WriteLine(txtFromRun.Text);
                            //    Console.WriteLine(txtFromRun.InnerText);
                            //}
                        }
                    }
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void ChangeSingleMergefield(string filePath, string mergefieldName = "DEBTOR__First_name_excl_middle", string correctValue = "Harry")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";

                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(completeMergeFieldName)))
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    var runToBeChanged = paragraph.Descendants<Run>().ToList()[3];
                    var elementToBeChanged = runToBeChanged.InnerText;
                    Console.WriteLine("Before: " + runToBeChanged.InnerText);
                    var textToBeChanged = runToBeChanged.Descendants<Text>().FirstOrDefault();
                    textToBeChanged.Text = correctValue;
                    Console.WriteLine("After: " + runToBeChanged.InnerText);
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void ChangeEmptyMergefield(string filePath, string mergefieldName = "DEBTOR__Middle_name", string correctValue = "Harry")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";

                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(completeMergeFieldName)))
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    var clone = paragraph.Descendants<Run>().FirstOrDefault().Clone();
                    paragraph.AppendChild((Run)clone);
                    var clone2 = paragraph.Descendants<Run>().FirstOrDefault().Clone();
                    paragraph.AppendChild((Run)clone2);

                    var changeFldCharType = paragraph.Descendants<Run>().ToList()[2];
                    // changeFldCharType.LastChild.FieldCharType = "separate";
                    var thirdRunChildren = changeFldCharType.ChildElements.ToList();
                    var fldTypeElement = changeFldCharType.Descendants().Last();
                    // fldTypeElement.FieldCharType = "separate";

                    var runToBeChanged = paragraph.Descendants<Run>().ToList()[3];
                    runToBeChanged.RsidRunProperties = "007F440F";
                    runToBeChanged.RsidRunAddition = "005A020C";
                    runToBeChanged.AppendChild(new Text(correctValue));
                }
                // doc.MainDocumentPart.Document.Save();
            }
        }

        public void ReadLeapForm(string myFilePath)
        {
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
