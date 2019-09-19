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
            MyMethodContainer methodContainer = new MyMethodContainer(basePath);
            Person person = new Person { FirstName = "Robert", MiddleName = "Harold", LastName = "Williamson" };
            string[] fieldNames = new string[3] { "DEBTOR__First_name_excl_middle", "DEBTOR__Middle_name", "DEBTOR__People_Last_Name" };
            string[] personDetails = new string[3] { person.FirstName, person.MiddleName, person.LastName };

            // ExperimentXML experiment = new ExperimentXML();
            // experiment.WriteToWordDoc($"{basePath}ExperimentDoc.docx", "Hello world");
            // Console.WriteLine(experiment.CreateWordDoc($"{basePath}Create{DateTime.Now.Millisecond}.zip", "Hello world!"));

            // reader.FindAllIfParagraphs(mainFilePath);
            // var fields = reader.ReadLeapForm($"{basePath}{fileName}.docx");
            // DictionaryBuilder builder = new DictionaryBuilder(fields);
            // var values = builder.GetMergefieldDictionary();
            // reader.GetFilerType(fields);

            // methodContainer.RemoveHeaderIfParagraphs(mainFilePath);
            // methodContainer.CleanMergefieldV2(mainFilePath, "DEBTOR__First_name_excl_middle", "Harold");
            // methodContainer.ExchangeTextValue(mainFilePath, "6789", "0000");
            // methodContainer.ChangeSingleMergefield(mainFilePath, "DEBTOR2__Middle_name", "Robert");
            // methodContainer.ChangeEmptyMergefield(mainFilePath, "BANKRUPTCY_DE__Case_number", "12345");
            // methodContainer.CheckCheckbox(mainFilePath, "Chapter 7");
            // methodContainer.UncheckCheckbox(mainFilePath, "Chapter 7");
            // methodContainer.SaveZipFile($"{basePath}{fileName}.docx", fileName);

            reader.ReadSmokeballForm(smokeballFilePath);

            Console.WriteLine("The program is complete.");
            Console.ReadLine();
        }
    }
}
