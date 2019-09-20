using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Spire.Doc.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExtractMergeFields
{
    class Program
    {
        private static readonly string basePath = "../../Docs/";
        private static readonly string mainLeapFilePath = $"{basePath}LeapBankruptcyPetition.docx";
        private static readonly string cloneLeapFilePath = $"{basePath}BankruptcyPetitionClone.docx";
        private static string leapFileName = "LeapBankruptcyPetition";
        private static readonly string smokeballFilePath = $"{basePath}SmokeballBankruptcyPetition.docx";
        static void Main()
        {
            Reader reader = new Reader(basePath);
            LeapDocCleaner cleaner = new LeapDocCleaner(mainLeapFilePath);
            MyMethodContainer methodContainer = new MyMethodContainer(basePath, new DocumentService());
            Person person = new Person { FirstName = "Robert", MiddleName = "Harold", LastName = "Williamson" };
            string[] fieldNames = new string[3] { "DEBTOR__First_name_excl_middle", "DEBTOR__Middle_name", "DEBTOR__People_Last_Name" };
            string[] personDetails = new string[3] { person.FirstName, person.MiddleName, person.LastName };

            string location = methodContainer.CopyFileToNewLocation(mainLeapFilePath);
            //cleaner.RemoveHeaderIfParagraphs(location, "DEBTOR__First_name_excl_middle");
            //cleaner.CleanMergefieldV2(location, "DEBTOR__First_name_excl_middle");

            //foreach(var fieldName in fieldNames)
            //{
            //    cleaner.RemoveHeaderIfParagraphs(location, fieldName);
            //    cleaner.CleanMergefieldV2(location, fieldName);
            //}

            for(int i = 0; i< fieldNames.Length; i++)
            {
                cleaner.RemoveHeaderIfParagraphs(location, fieldNames[i]);
                cleaner.CleanMergefieldV2(location, fieldNames[i], personDetails[i]);
            }

            // methodContainer.FixTextValue(location, "0000", "6789");
            // methodContainer.ChangeSingleMergefield(location, "DEBTOR2__Middle_name", "Robert");
            // methodContainer.ChangeEmptyMergefield(location, "BANKRUPTCY_DE__Case_number", "12345");
            // methodContainer.CheckCheckbox(location, "Chapter 7");
            // methodContainer.UncheckCheckbox(location, "Chapter 7");
            // methodContainer.SaveZipFile($"{basePath}{fileName}.docx", fileName);

            // leap dictionary

            //var leapFields = reader.ReadLeapForm($"{basePath}{leapFileName}.docx");
            //DictionaryBuilder builder = new DictionaryBuilder(leapFields);
            //var values = builder.GetMergefieldDictionary();

            // smokeball dictionary

            //var smokeballFields = reader.ReadSmokeballForm(smokeballFilePath);
            //var fieldDictionary = reader.ConvertFieldDataToDictionary(smokeballFields);
            //foreach (var k in fieldDictionary.Keys)
            //{
            //    Console.WriteLine($"{k}: {fieldDictionary[k]}");
            //}

            // reader.GetFilerType(leapFields);
            // reader.FindAllIfParagraphs(mainFilePath);

            // Console.WriteLine(reader.GetPageCount(mainLeapFilePath));
            // Console.WriteLine(reader.GetPageCount(smokeballFilePath));

            // reader.ReadSmokeballForm(smokeballFilePath);
            // reader.ReadLeapForm(mainLeapFilePath);

            Console.WriteLine("The program is complete.");
            Console.ReadLine();
        }
    }
}
