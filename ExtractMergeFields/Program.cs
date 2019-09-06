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
        private static readonly string mainFilePath = $"{basePath}LeapBankruptcyPetition.docx";
        static void Main()
        {
            MyMethodContainer methodContainer = new MyMethodContainer(basePath);
            Person person = new Person { FirstName = "Robert", MiddleName = "Harold", LastName = "Williamson" };
            string[] fieldNames = new string[3] { "DEBTOR__First_name_excl_middle", "DEBTOR__Middle_name", "DEBTOR__People_Last_Name" };
            string[] personDetails = new string[3] { person.FirstName, person.MiddleName, person.LastName };

            // ExperimentXML experiment = new ExperimentXML();
            // experiment.WriteToWordDoc($"{basePath}ExperimentDoc.docx", "Hello world");
            // Console.WriteLine(experiment.CreateWordDoc($"{basePath}Create{DateTime.Now.Millisecond}.zip", "Hello world!"));

            methodContainer.ChangeEmptyMergefield(mainFilePath);

            Console.ReadLine();
        }
    }
}
