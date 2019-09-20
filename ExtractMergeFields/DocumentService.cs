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
                }
            }
        }
        public void ExecuteTextChange(Paragraph paragraph, string correctValue)
        {
            var runToBeChanged = paragraph.Descendants<Run>().ToList()[3];
            var textToBeChanged = runToBeChanged.Descendants<Text>().FirstOrDefault();
            textToBeChanged.Text = correctValue;
        }
    }
}
