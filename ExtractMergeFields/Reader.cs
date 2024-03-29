﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public class Reader
    {
        readonly string _basePath;
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

        public IEnumerable<Spire.Doc.Fields.Field> ReadLeapForm(string myFilePath)
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

            //foreach (var f in mergeFields)
            //{
            //    Console.Write($"{f.Code.Substring(12, f.Code.Length - 12)}, {f.Text}\n");
            //}
            return mergeFields;
        }

        public IEnumerable<Spire.Doc.Fields.Field> ReadSmokeballForm(string filePath)
        {
            // Use spire to get a collection of fields in the smokeball document
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile($"{filePath}");

            var merge = document.MailMerge;
            FieldCollection fields = document.Fields;
            var mergeFields = from Spire.Doc.Fields.Field f in fields
                              select f;
            mergeFields = mergeFields.Where(x => !x.Code.Contains("CHECKBOX") && !x.Code.Contains("AUTOMATIONIF") && !x.Code.Contains("AUTOMATIONELSE") && !x.Code.Contains("AUTOMATIONENDIF") && !x.Code.Contains("AUTOMATIONENDELSE")).ToList();

            // To see all mergefields in the document, uncomment the following lines
            foreach (var f in mergeFields)
            {
                int matchIndex = GetMatchIndex(f);
                Console.WriteLine($"{f.Code.Substring(matchIndex)}");
            }

            var cleanMergefields = new List<Spire.Doc.Fields.Field>();

            foreach (var f in mergeFields)
            {
                if (f.FieldText != string.Empty && !cleanMergefields.Contains(f))
                {
                    cleanMergefields.Add(f);
                }
            }

            foreach(var f in cleanMergefields)
            {
                var mergeFieldNameStartIndex = GetMatchIndex(f);
            }
            return cleanMergefields;
        }

        public Dictionary<string, string> ConvertFieldDataToDictionary(IEnumerable<Spire.Doc.Fields.Field> fields)
        {
            Dictionary<string, string> fieldValuePairs = new Dictionary<string, string>();
            foreach(var f in fields)
            {
                var mergeFieldNameStartIndex = GetMatchIndex(f);
                string key = f.Code.Substring(mergeFieldNameStartIndex);
                if (!fieldValuePairs.ContainsKey(key))
                {
                    string value = f.FieldText;
                    fieldValuePairs.Add(key, value);
                }
            }
            return fieldValuePairs;
        }

        private int GetMatchIndex(Spire.Doc.Fields.Field field)
        {
            var match = Regex.Match(field.Code, @"\\\*\\\*");
            return match.Index + 4;
        }

        public void GetFilerType(IEnumerable<Spire.Doc.Fields.Field> fields)
        {
            List<Spire.Doc.Fields.Field> mergeFieldMap = new List<Spire.Doc.Fields.Field>();
            foreach (Spire.Doc.Fields.Field field in fields)
            {
                string mergefieldName = field.Code.Substring(12, field.Code.Length - 12);
                string value = field.Text;
                if (mergefieldName == "BANKRUPTCY_DE__Type_of_debtor ")
                {
                    mergeFieldMap.Add(field);
                    Console.WriteLine($"{{ \"{mergefieldName}\", \"{value}\" }},");
                }
            }
            Console.WriteLine($"There are {mergeFieldMap.Count()} items in the list.");
            return;
        }

        public int GetPageCount(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                // This number can be inaccurate since it is only the result of the last time Word calculated it
                return int.Parse(doc.ExtendedFilePropertiesPart.Properties.Pages.Text);
            }
        }
    }
}
