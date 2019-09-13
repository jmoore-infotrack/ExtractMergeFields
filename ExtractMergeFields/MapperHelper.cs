using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractMergeFields
{
    public static class MapperHelper
    {
        public static List<KeyValuePair<string, string>> BogusValues =
            new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string> ( "DEBTOR2__First_name_excl_middle ", "Benjamin" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Middle_name ", "Elvin" ),
                new KeyValuePair<string, string> ( "DEBTOR2__People_Last_Name ", "Wasser" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_firstname_1 ", "Frederick" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_middlename_1 ", "A." ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_lastname_1 ", "Lobos" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_firstname_2 ", "Jermaine" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_middlename_2 ", "L." ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_lastname_2 ", "Edmonds" ),
                new KeyValuePair<string, string> ( "DEBTOR2__SSN_last_4_digits ", "1254" ),
                new KeyValuePair<string, string> ( "EBTOR__First_name_excl_middle ", "Susan" ),
                new KeyValuePair<string, string> ( "EBTOR2__First_name_excl_middle ", "Benjamin" ),
                new KeyValuePair<string, string> ( "DEBTOR__Middle_name ", "Beth" ),
                new KeyValuePair<string, string> ( "EBTOR2__Middle_name", "Elvin" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_EIN_1 ", "55-63245" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_name_1 ", "Mortor & Co." ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_name_2 ", "Excelsior Music Corporation" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_EIN_1_fi ", "55" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_EIN_1_la ", "5-63245" ),
                new KeyValuePair<string, string> ( "BANKRUPTCY_DE__Type_of_debtor ", "Joint (Debtor 1 and Debtor 2)" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Other_business_EIN_2_la ", "6-12342" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_Number ", "45" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_Street ", "Baldwin Street" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Street_address_line_1 ", "45 Baldwin Street" ),
                new KeyValuePair<string, string> ( "DEBTOR__Street_address_line_1 ", "103 Des Moines Ct" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_City ", "Saddle River" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_State ", "NJ" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_Zip ", "07458" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Address_County ", "Bergen" ),
                new KeyValuePair<string, string> ( "DEBTOR2__PO_Box_Number ", "6534" ),
                new KeyValuePair<string, string> ( "DEBTOR2__PO_Box_City ", "Tacoma" ),
                new KeyValuePair<string, string> ( "DEBTOR2__PO_Box_State ", "WA" ),
                new KeyValuePair<string, string> ( "DEBTOR2__PO_Box_Zip ", "98466" ),
                new KeyValuePair<string, string> ( "DEBTOR__Card_count ", "2" ),
                new KeyValuePair<string, string> ( "BANKRUPTCY_DE__Venue ", "Over 180 days" ),
                new KeyValuePair<string, string> ( "DEBTOR__Is_sole_proprietor ", "False" ),
                new KeyValuePair<string, string> ( "DEBTOR__Name_of_business ", "SUP Bro" ),
                new KeyValuePair<string, string> ( "DEBTOR__Business_street_add1 ", "Paddle Street" ),
                new KeyValuePair<string, string> ( "DEBTOR__Business_city ", "Atlanta" ),
                new KeyValuePair<string, string> ( "DEBTOR__Business_state ", "GA" ),
                new KeyValuePair<string, string> ( "DEBTOR__Business_zip ", "58346" ),
                new KeyValuePair<string, string> ( "DEBTOR__Is_hazard ", "False" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_description ", "Fire" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_attentions ", "There is a fire" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_property_number ", "48" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_property_street1 ", "Renfrew Street" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_property_city ", "Trenton" ),
                new KeyValuePair<string, string> ( "DEBTOR__Hazard_property_state ", "NJ" ),
                new KeyValuePair<string, string> ( "BANKRUPTCY_DE__Other_debt ", "different debt" ),
                new KeyValuePair<string, string> ( "DEBTOR__Full_name ", "Rory J. Gilmore" ),
                new KeyValuePair<string, string> ( "BANKRUPTCY_DE__Electronic_signature ", "True" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Full_name ", "Doris Day Belmont" ),
                new KeyValuePair<string, string> ( "MATTER__Person_Acting_full_name ", "Janeth Nikodem" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Firm_name ", "Hanna, Zhu & Cassar Lawyers" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Street_number ", "207" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Street_name ", "Kent Street" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Unit_suite_or_office ", "2" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__City ", "Sydney" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__State ", "MI" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Zip ", "2001" ),
                new KeyValuePair<string, string> ( "FIRM_DETAILS__Phone_formatted ", "(028) 273-7500220" ),
                new KeyValuePair<string, string> ( "DEBTOR__Phone_number_formatted ", "() -" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Phone_number_formatted ", "(201) 565-4141" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Card_E_Mail_Addresses ", "Fran@email" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Full_name ", "Benjamin Elvin Wasser" ),
                new KeyValuePair<string, string> ( "DEBTOR2__Full_name ", "Fran Brooke Collins" )
            };
    }
}
