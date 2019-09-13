using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public class DictionaryBuilder
    {
        IEnumerable<Spire.Doc.Fields.Field> _fields;

        public DictionaryBuilder(IEnumerable<Spire.Doc.Fields.Field> fields)
        {
            _fields = fields;
        }

        public Dictionary<string,string> GetMergefieldDictionary()
        {
            Dictionary<string, string> mergeFieldMap = new Dictionary<string, string>();
            foreach(Spire.Doc.Fields.Field field in _fields)
            {
                string mergefieldName = field.Code.Substring(12, field.Code.Length - 12);
                string value = field.Text;
                if (!mergeFieldMap.Keys.Contains(mergefieldName) && (value != "") && !MapperHelper.BogusValues.Any(kvp => kvp.Key == mergefieldName && kvp.Value == value))
                {
                    mergeFieldMap.Add(mergefieldName, value);
                    Console.WriteLine($"{{ \"{mergefieldName}\", \"{value}\" }},");
                }
            }
            Console.WriteLine($"There are {mergeFieldMap.Count()} items in the dictionary.");
            return mergeFieldMap;
        }
    }
}
