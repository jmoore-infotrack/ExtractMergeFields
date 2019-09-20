using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc.Collections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public class MyMethodContainer
    {
        string _basePath;
        IDocumentService _service;
        public MyMethodContainer(string basePath, IDocumentService service)
        {
            _basePath = basePath;
            _service = service;
        }

        public void FixTextValue(string filePath, string incorrectValue = "Harry", string correctValue = "Harold")
        {
            // Updates an incorrect value of a complex mergefield with a correct value
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var runs = doc.MainDocumentPart.Document.Descendants<Run>();
                _service.FixIncorrectTextValue(runs, incorrectValue, correctValue);
                doc.Save();
            }
        }

        public void ChangeSingleMergefield(string filePath, string mergefieldName = "DEBTOR__First_name_excl_middle", string correctValue = "Robert")
        {
            // Updates a merge field's value when we know the field name but not the current value
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string completeMergeFieldName = $" MERGEFIELD {mergefieldName} ";
                var fields = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().Where(x => x.Text.Contains(completeMergeFieldName));

                _service.ChangeSingleMergefield(fields, correctValue);

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

                _service.InputValueInEmptyMergefield(fieldsWithMergefieldName, correctValue);

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
            // since AutoSave is true by default, Close and Dispose saves changes
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

        public void SaveZipFile(string openPath, string fileName)
        {
            // saves a document into a zip collection that exposes its XML
            using (WordprocessingDocument doc = WordprocessingDocument.Open(openPath, true))
            {
                string savePath = $"{_basePath}/{fileName}.zip";
                doc.SaveAs(savePath);
            }
        }

        public string CopyFileToNewLocation(string filePath)
        {
            int seconds = DateTime.Now.Second;
            string newLocation = filePath.Remove(filePath.Length - 5, 5) + seconds + ".docx";
            File.Copy(filePath, newLocation);

            return newLocation;
        }
    }
}
