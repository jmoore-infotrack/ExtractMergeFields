﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
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

        public void ExchangeTextValue(string filePath, string incorrectValue = "Harry", string correctValue = "Harold")
        {
            // Updates an incorrect value of a complex mergefield with a correct value
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
            // not currently usable
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

        public void ChangeSingleMergefield(string filePath, string mergefieldName = "DEBTOR__First_name_excl_middle", string correctValue = "Robert")
        {
            // Updates a merge field's value when we know the field name but not the current value
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";

                foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(completeMergeFieldName)))
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    if (paragraph.Count() == 5 || paragraph.Count() == 6)
                    {
                        ExecuteTextChange(field, paragraph, correctValue);
                    }
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void ChangeEmptyMergefield(string filePath, string mergefieldName = "DEBTOR__Middle_name", string correctValue = "Leonard")
        {
            // should update the field's value when it is currently empty, while preserving the mergefield
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";

                var fieldsWithMergefieldName = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Equals(completeMergeFieldName));
                Console.WriteLine($"There are {fieldsWithMergefieldName.Count()} fields with the following MergefieldName: {mergefieldName}");
                foreach (FieldCode field in fieldsWithMergefieldName)
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    Run mergefieldRunElement = (Run)field.Parent;
                    Run precedingElement = (Run)mergefieldRunElement.PreviousSibling();
                    if (DiscoverIf(mergefieldRunElement))
                    {
                        continue;
                    }
                    // An empty paragraph can have 3 or 4 ChildElements, depending on whether there is a properties element
                    if(paragraph.Count() == 3 || paragraph.Count() == 4)
                    {
                        var clone = precedingElement.Clone();
                        var clone2 = precedingElement.Clone();

                        paragraph.InsertAt((Run)clone, paragraph.ChildElements.Count() - 1);
                        paragraph.InsertAt((Run)clone2, paragraph.ChildElements.Count() - 1);

                        // This changes the middle element's field type to "separate," a necessary component of the mergefield structure
                        var runs = paragraph.Descendants<Run>().ToList();
                        var separateRun = runs[2];
                        var fldChar = separateRun.Descendants<FieldChar>().FirstOrDefault();
                        fldChar.FieldCharType.Value = FieldCharValues.Separate;
                        fldChar.FieldLock = null;

                        // This updates the second to last element and inserts the desired value
                        var runToBeChanged = runs[3];
                        runToBeChanged.RsidRunProperties = "007F440F";
                        runToBeChanged.RsidRunAddition = "005A020C";
                        runToBeChanged.RemoveChild(runToBeChanged.ChildElements[1]);
                        runToBeChanged.AppendChild(new Text(correctValue));
                    }
                    else
                    {
                        Console.WriteLine("There is already data in the fields.");
                        continue;
                    }
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void CheckCheckbox(string filePath, string locatorText)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var chapterSevenTextNode = doc.MainDocumentPart.RootElement.Descendants<Text>().Where(x => x.Text.Equals(locatorText)).FirstOrDefault();
                var paragraph = chapterSevenTextNode.Parent.Parent;

                // First we set the checked value to 1
                SdtRun checkerElement = paragraph.Descendants<SdtRun>().FirstOrDefault();
                SdtProperties properties = checkerElement.Descendants<SdtProperties>().FirstOrDefault();
                SdtContentCheckBox checkbox = (SdtContentCheckBox)properties.ChildElements[2];
                checkbox.Checked.Val = OnOffValues.One;

                // Next we change the symbol shown
                var content = properties.NextSibling().NextSibling();
                SymbolChar symbol = content.Descendants<SymbolChar>().FirstOrDefault();
                symbol.Char = "F052";
                symbol.Font = "Wingdings 2";
            }
        }

        public void UncheckCheckbox(string filePath, string locatorText)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var textNode = doc.MainDocumentPart.RootElement.Descendants<Text>().Where(x => x.Text.Equals(locatorText)).FirstOrDefault();
                var paragraph = textNode.Parent.Parent;

                // First we set the checked value to 0
                SdtRun checkerElement = paragraph.Descendants<SdtRun>().FirstOrDefault();
                SdtProperties properties = checkerElement.Descendants<SdtProperties>().FirstOrDefault();
                SdtContentCheckBox checkbox = (SdtContentCheckBox)properties.ChildElements[2];
                checkbox.Checked.Val = OnOffValues.Zero;

                // Next we change the symbol shown
                var content = properties.NextSibling().NextSibling();
                SymbolChar symbol = content.Descendants<SymbolChar>().FirstOrDefault();
                symbol.Char = "F072";
                symbol.Font = "Wingdings";
            }
        }

        public void ExecuteTextChange(FieldCode field, Paragraph paragraph, string correctValue)
        {
            var runToBeChanged = paragraph.Descendants<Run>().ToList()[3];
            Console.WriteLine("Before: " + runToBeChanged.InnerText);
            var textToBeChanged = runToBeChanged.Descendants<Text>().FirstOrDefault();
            textToBeChanged.Text = correctValue;
            Console.WriteLine("After: " + runToBeChanged.InnerText);
        }

        public bool DiscoverIf(Run run)
        {
            bool ifExists = false;
            var siblings = run.Parent.ChildElements.ToList();
            foreach (var s in siblings)
            {
                if (s.InnerText == " IF ")
                {
                    ifExists = true;
                    break;
                }
            }
            return ifExists;
        }

        public void SaveZipFile(string openPath, string fileName)
        {
            // saves a document into a zip collection that exposes its XML
            using (WordprocessingDocument doc = WordprocessingDocument.Open(openPath, true))
            {
                string savePath = $"{_basePath}/{fileName}.zip";
                doc.SaveAs(savePath);
            }
        }

        public void CleanMergefield(string filePath, string mergefieldName = "DEBTOR2__Middle_name", string correctValue = "")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $"MERGEFIELD {mergefieldName} ";

                var fieldsWithMergefieldName = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(mergefieldName));
                foreach (FieldCode field in fieldsWithMergefieldName)
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    if (paragraph.Count() <= 6) return;
                    var firstBeginDone = false;
                    var firstMergefieldDone = false;
                    var firstSeparateDone = false;
                    var firstTextDone = false;
                    var firstEndDone = false;

                    var runs = paragraph.Descendants<Run>().ToList();
                    foreach (Run run in runs)
                    {
                        if(run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.Begin) && firstBeginDone != true)
                        {
                            firstBeginDone = true;
                            continue;
                        }
                        if (run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.Separate) && firstSeparateDone != true)
                        {
                            firstSeparateDone = true;
                            continue;
                        }
                        if (run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.End) && firstEndDone != true)
                        {
                            firstEndDone = true;
                            continue;
                        }
                        if (run.Descendants<FieldCode>().Any(x => x.Text.Contains(mergefieldName)) && firstMergefieldDone != true)
                        {
                            firstMergefieldDone = true;
                            continue;
                        }
                        if (run.RsidRunProperties != null && run.RsidRunAddition != null && firstTextDone != true)
                        {
                            FieldCode fieldcode = (FieldCode)run.ChildElements[1];
                            fieldcode.Text = "value";
                            firstTextDone = true;
                            continue;
                        }
                        run.Remove();
                    }
                    Run mergeFieldElement = (Run)paragraph.ChildElements[4].Clone();
                    paragraph.InsertAt(mergeFieldElement, 1);
                    paragraph.ChildElements[5].Remove();
                }
                doc.Save();
            }
        }
    }
}
