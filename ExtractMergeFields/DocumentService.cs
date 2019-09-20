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
    public class DocumentService : IDocumentService
    {
        public void FixIncorrectTextValue(IEnumerable<Run> runs, string incorrectValue, string correctValue)
        {
            foreach (Run run in runs)
            {
                foreach (Text txtFromRun in run.Descendants<Text>().Where(a => a.Text == incorrectValue))
                {
                    txtFromRun.Text = correctValue;
                }
            }
        }

        public void ChangeSingleMergefield(IEnumerable<FieldCode> fields, string correctValue)
        {
            foreach (FieldCode field in fields)
            {
                var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                if (paragraph.Count() == 5 || paragraph.Count() == 6)
                {
                    ExecuteTextChange(paragraph, correctValue);
                    return;
                }
            }
        }

        public void ExecuteTextChange(Paragraph paragraph, string correctValue)
        {
            var runToBeChanged = paragraph.Descendants<Run>().ToList()[3];
            var textToBeChanged = runToBeChanged.Descendants<Text>().FirstOrDefault();
            textToBeChanged.Text = correctValue;
        }

        public void InputValueInEmptyMergefield(IEnumerable<FieldCode> fields, string correctValue)
        {
            foreach (FieldCode field in fields)
            {
                var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
                Run mergefieldRunElement = (Run)field.Parent;
                Run precedingElement = (Run)mergefieldRunElement.PreviousSibling();
                if (DiscoverIf(mergefieldRunElement))
                {
                    continue;
                }
                // An empty paragraph can have 3 or 4 ChildElements, depending on whether there is a properties element
                if (paragraph.Count() == 3 || paragraph.Count() == 4)
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
    }
}
