using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExtractMergeFields
{
    class LeapDocCleaner
    {
        readonly string _leapFilePath;
        public LeapDocCleaner(string leapFilePath) {
            _leapFilePath = leapFilePath;
        }
        public void CleanMergefieldV2(string filePath, string mergefieldName = "DEBTOR__Middle_name", string correctValue = "test")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $"MERGEFIELD {mergefieldName} ";

                var fieldsWithMergefieldName = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(mergefieldName));
                foreach (FieldCode field in fieldsWithMergefieldName)
                {
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    if (paragraph.Count() <= 6) continue;
                    var firstBeginDone = false;
                    var firstMergefieldDone = false;
                    var firstSeparateDone = false;
                    var firstTextDone = false;
                    var firstEndDone = false;


                    var runs = paragraph.Descendants<Run>().ToList();
                    var tempRuns = new Run[5];
                    foreach (Run run in runs)
                    {
                        if (run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.Begin) && firstBeginDone != true)
                        {
                            tempRuns[0] = run;
                            firstBeginDone = true;
                        }
                        else if (run.Descendants<FieldCode>().Any(x => x.Text.Contains(mergefieldName)) && firstMergefieldDone != true)
                        {
                            tempRuns[1] = run;
                            firstMergefieldDone = true;
                        }
                        else if (run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.Separate) && firstSeparateDone != true)
                        {
                            tempRuns[2] = run;
                            firstSeparateDone = true;
                        }
                        else if (run.RsidRunProperties != null && run.RsidRunAddition != null && firstTextDone != true)
                        {
                            FieldCode fieldcode = (FieldCode)run.ChildElements[1];
                            fieldcode.Text = correctValue;
                            tempRuns[3] = run;
                            firstTextDone = true;
                        }
                        else if (run.Descendants<FieldChar>().Any(x => x.FieldCharType == FieldCharValues.End) && firstEndDone != true)
                        {
                            tempRuns[4] = run;
                            firstEndDone = true;
                        }
                        run.Remove();
                    }
                    foreach (Run run in tempRuns)
                    {
                        paragraph.AppendChild(run);
                    }
                }
                doc.Save();
            }
        }

        public void RemoveHeaderIfParagraphs(string filePath, string mergefieldName = "DEBTOR__Middle_name")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $"MERGEFIELD {mergefieldName} ";

                var fieldsWithMergefieldName = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(mergefieldName));

                Console.WriteLine($"There are {fieldsWithMergefieldName.Count()} fields in the document.");
                int incrementer = 0;

                foreach (FieldCode field in fieldsWithMergefieldName)
                {
                    incrementer++;
                    Console.WriteLine($"We have checked the fields {incrementer} times.");
                    var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                    var previousParagraph = paragraph?.PreviousSibling();
                    if (previousParagraph == null) continue;
                    if (previousParagraph.Descendants<FieldCode>().Any(x => x.Text.Contains(mergefieldName)))
                    {
                        paragraph.PreviousSibling().Remove();
                    }
                }
                //int seconds = DateTime.Now.Second;
                //doc.SaveAs(filePath.Remove(filePath.Length - 5, 5) + seconds + ".docx");
                doc.Save();
            }
        }
    }
}
