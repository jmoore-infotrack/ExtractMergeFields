using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExtractMergeFields
{
    public interface IDocumentService
    {
        void FixIncorrectTextValue(IEnumerable<Run> runs, string incorrectValue, string correctValue);

        void ChangeSingleMergefield(IEnumerable<FieldCode> fields, string correctValue);

        void ExecuteTextChange(Paragraph paragraph, string correctValue);

        void InputValueInEmptyMergefield(IEnumerable<FieldCode> fields, string correctValue);
    }
}