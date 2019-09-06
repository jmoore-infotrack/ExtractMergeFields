using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using Spire.Doc.Reporting;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExtractMergeFields
{
    class Program
    {
        private static readonly string basePath = "../../Docs/";
        private static readonly string mainFilePath = $"{basePath}LeapBankruptcyPetition.docx";
        static void Main()
        {
            Person person = new Person { FirstName = "Robert", MiddleName = "Harold", LastName = "Williamson" };
            string[] fieldNames = new string[3] { "DEBTOR__First_name_excl_middle", "DEBTOR__Middle_name", "DEBTOR__People_Last_Name" };
            string[] personDetails = new string[3] { person.FirstName, person.MiddleName, person.LastName };
            
            // ExperimentXML experiment = new ExperimentXML();
            // experiment.WriteToWordDoc($"{basePath}ExperimentDoc.docx", "Hello world");
            // Console.WriteLine(experiment.CreateWordDoc($"{basePath}Create{DateTime.Now.Millisecond}.zip", "Hello world!"));

            // ExamineEstimateOfFiguresLetter();

            // ReadSmokeballForm();

            // ReadLeapForm(mainFilePath);

            // InsertValueWithSpire(fieldNames, personDetails);

            // InsertValueWithInterop(fieldNames, personDetails);

            // MailMergeWithOpenXML(mainFilePath);

            ChangeSingleMergefieldOpenXML(mainFilePath);

            // ChangeTextValueOpenXML(mainFilePath);

            // SaveZipFile(mainFilePath);

            Console.ReadLine();
        }

        private static void SaveZipFile(string openPath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(openPath, true))
            {
                string savePath = $"{basePath}Petition.zip";
                doc.SaveAs(savePath);
            }
        }

        private static void ChangeTextValueOpenXML(string filePath, string incorrectValue = "Harold", string correctValue = "Harry")
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

        private static void MailMergeWithOpenXML(string filePath, string incorrectValue = "Harold", string correctValue = "Harry")
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

        private static void ChangeSingleMergefieldOpenXML(string filePath, string mergefieldName = "DEBTOR__First_name_excl_middle", string correctValue = "Harry")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";

                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(completeMergeFieldName)))
                {
                        var paragraph = field.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
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

        private static void InsertValueWithInterop(string[] fieldNames, string[] personDetails)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            object fileName = @"C:\Users\Justin.Moore\source\learning\ExtractMergeFields\ExtractMergeFields\Docs\LeapBankruptcyPetition.docx";
            object saveFileName = @"C:\Users\Justin.Moore\source\learning\ExtractMergeFields\ExtractMergeFields\Docs\LeapBankruptcyPetitionNew.docx";
            // object fileName = $"{basePath}LeapBankruptcyPetition.docx";

            Application oWord = new Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            object isVisible = false;

            if (File.Exists((string)fileName))
            {
                wordDoc = oWord.Documents.Open(ref fileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                wordDoc.Activate();

                string pMergeField = fieldNames[0];
                string pValue = personDetails[0];

                foreach (Microsoft.Office.Interop.Word.Field myMergeField in wordDoc.Fields)
                {
                    Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;
                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        String fieldName = fieldText.Substring(11, fieldText.Length - 11);
                        fieldName = fieldName.Trim();
                        if (fieldName == pMergeField)
                        {
                            myMergeField.Select();
                            oWord.Selection.TypeText(pValue);
                        }
                    }
                }
                wordDoc.SaveAs(ref saveFileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                oWord.Quit();
            }
        }

        private static void InsertValueWithSpire(string[] fieldNames, string[] personDetails)
        {
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile($"{basePath}LeapBankruptcyPetition.docx");

            document.MailMerge.Execute(fieldNames, personDetails);
            document.SaveToFile($"{basePath}LeapBankruptcyPetitionModified.docx");

            // this seems to dispose of itself
        }

        private static void ReadLeapForm(string myFilePath)
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

        private static void ReadSmokeballForm()
        {
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile($"{basePath}SmokeballBankruptcyPetition.docx");

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

        private static void ExamineEstimateOfFiguresLetter()
        {
            Spire.Doc.Document estimateLetter = new Spire.Doc.Document();
            estimateLetter.LoadFromFile($"{basePath}EstimateOfFiguresLetter.docx");

            var merge = estimateLetter.MailMerge;
            FieldCollection fields = estimateLetter.Fields;

            var mergeFields = from Spire.Doc.Fields.Field f in fields
                              where f.DocumentObjectType.ToString() != ""
                              select f;
            var fieldItems = mergeFields.ToList();

            foreach (var f in fieldItems)
            {
                Console.WriteLine($"{f.Code}");
                Console.WriteLine($"{f.Value}, {f.FieldText}\n\n");
            }
        }
    }
}
