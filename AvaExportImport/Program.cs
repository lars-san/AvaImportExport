using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualBasic.FileIO;//.TextFieldParser;

class Program
{
    static void Main()
    {
        //Command entry
        bool exit = false;
        string ProcessCode = "";
        string CompanyCode = "";
        bool taxincludes = false;
        ConsoleKeyInfo keypress; // ready the variable used for accepting user input

        Console.WriteLine("Instructions");
        Console.WriteLine("This program will take export.csv in the local directory and write import.csv in the local directory, reformatting it.");
        Console.WriteLine("Press \"P\" to specify the process code.");
        Console.WriteLine("Press \"C\" to specify the company code.");
        Console.WriteLine("Press \"K\" to switch to tax includes for all records. Yes, I said \"K\".");
        Console.WriteLine("Press \"I\" to create the import.csv file.");
        Console.WriteLine("Press \"0\" to take excel.csv and write 0's to zip codes where needed to fixed.csv.");
        Console.WriteLine("");
        Console.WriteLine("You can press Esc to exit.");
        while (!exit) // This is the main loop and will end when the user presses the Esc key.
        {
            keypress = Console.ReadKey(true); //This is listed with true so the key press is not shown.
            if (keypress.Key == ConsoleKey.I)
            {
                ConvertExport(ProcessCode, CompanyCode, taxincludes);
            }
            if (keypress.Key == ConsoleKey.P)
            {
                Console.Write("Enter the desired process code: ");
                ProcessCode = Console.ReadLine();
                if (ProcessCode == "")
                { Console.WriteLine("Process code \"4\" will be used."); }
            }
            if (keypress.Key == ConsoleKey.C)
            {
                Console.WriteLine("Enter the desired company code:");
                Console.WriteLine("(Enter \"[Blank]\" for no specified company code)");
                CompanyCode = Console.ReadLine();
                if (CompanyCode == "")
                { Console.WriteLine("The company code in the export.csv file will be used."); }
                if (CompanyCode == "[Blank]")
                { Console.WriteLine("A blank will be entered for the company code."); }
            }
            if (keypress.Key == ConsoleKey.K)
            {
                Console.WriteLine("Tax includes will be used. Restart to reset this to normal behavior.");
                taxincludes = true;
            }
            if (keypress.Key == ConsoleKey.NumPad0 || keypress.Key == ConsoleKey.D0)
            {
                CorrectZip();
            }
            if (keypress.Key == ConsoleKey.Escape) // This checks to see if the key pressed was Esc
            { exit = true; } // This variable change will end the main loop and close the program.
        }
    }

    static void ConvertExport(string ProcessCode, string CompanyCode, bool taxincludes)
    {
        int lineCount = File.ReadAllLines("export.csv").Length;
        Console.WriteLine("Processing " + lineCount + " lines.");
        string[,] data = new string[lineCount, 76];
        TextFieldParser parser = new TextFieldParser("export.csv");

        parser.HasFieldsEnclosedInQuotes = true;
        parser.SetDelimiters(",");
        string[] fields;
        int row = 0;
        int column = 0;
        bool taxdateoverride = false;
        double line_amount;
        double tax_amount;
        string ti_string;
        while (!parser.EndOfData)
        {
            fields = parser.ReadFields();
            foreach (string field in fields)
            {
                //Console.WriteLine(field); // This can be used to see data as it loads into the data array
                //data[row, column] = field;
                switch (column)
                {
                    case 0:
                        if (ProcessCode != "")
                        { data[row, 0] = ProcessCode; }
                        else
                        { data[row, 0] = "4"; } // This will set the process code to 4.
                        break;
                    case 1:
                        if (CompanyCode == "")
                        { data[row, 4] = ""; } // case 1: is actually the CompanyID, which is not the company code. The document lines export doesn't appear to provide the company code
                        else
                        {
                            if (CompanyCode == "[Blank]")
                            { data[row, 4] = ""; }
                            else
                            { data[row, 4] = "\"" + CompanyCode + "\""; } //Why did this give an error?
                        }
                        break;
                    case 3:
                        data[row, 1] = "\"" + field + "\""; //This will need a change to allow for test imports
                        break;
                    case 4:
                        //data[row, 2] = "0"; //This is a temporary change, which should be rolled into an enhancement
                        if (field == "Sales Invoice")
                        { data[row, 2] = "1"; }
                        else if (field == "Return Invoice")
                        {
                            data[row, 2] = "5";
                            taxdateoverride = true;
                        }
                        else { data[row, 2] = field; }
                        break;
                    case 5:
                        data[row, 3] = field;
                        break;
                    case 6:
                        if (taxdateoverride)
                        {
                            data[row, 9] = field; // This will set the date override date to the date used in the original, but only for return invoices.
                            taxdateoverride = false;
                        }
                        break;
                    case 11:
                        data[row, 32] = field;
                        break;
                    case 12:
                        data[row, 5] = "\"" + field + "\"";
                        break;
                    case 13:
                        data[row, 6] = field;
                        break;
                    case 14:
                        data[row, 17] = field;
                        break;
                    case 15:
                        data[row, 31] = field;
                        break;
                    case 17:
                        data[row, 30] = field;
                        break;
                    case 18:
                        data[row, 29] = field;
                        break;
                    case 19:
                        data[row, 7] = field;
                        break;
                    case 20:
                        data[row, 10] = field;
                        break;
                    case 21:
                        data[row, 11] = "\"" + field + "\"";
                        break;
                    case 22:
                        data[row, 8] = field;
                        break;
                    case 25:
                        data[row, 12] = field;
                        break;
                    case 26:
                        data[row, 13] = field; // This will copy the line amount.
                        break;
                    case 27:
                        data[row, 14] = field;
                        break;
                    case 28:
                        data[row, 15] = field;
                        break;
                    case 29:
                        data[row, 16] = field;
                        break;
                    case 30:
                        data[row, 18] = field;
                        break;
                    case 34:
                        data[row, 41] = field; // This will copy the tax amount.
                        if (taxincludes && row != 0) // This looks for the tax includes flag, which can be set by the app user, and it skips the header row.
                        {
                            tax_amount = Convert.ToDouble(field);
                            line_amount = Convert.ToDouble(data[row, 13]);
                            line_amount = line_amount + tax_amount;
                            ti_string = Convert.ToString(line_amount);
                            data[row, 13] = ti_string;
                            data[row, 36] = "1";
                        }
                        break;
                    case 39:
                        data[row, 24] = "\"" + field + "\"";
                        break;
                    case 40:
                        data[row, 25] = "\"" + field + "\"";
                        break;
                    case 41:
                        data[row, 26] = field;
                        break;
                    case 42:
                        data[row, 28] = field;
                        break;
                    case 43:
                        data[row, 27] = "\"" + field + "\"";
                        break;
                    case 45:
                        data[row, 19] = "\"" + field + "\"";
                        break;
                    case 46:
                        data[row, 20] = "\"" + field + "\"";
                        break;
                    case 47:
                        data[row, 21] = field;
                        break;
                    case 48:
                        data[row, 23] = field;
                        break;
                    case 49:
                        data[row, 22] = "\"" + field + "\"";
                        break;
                }
                column = column + 1;
            }
            column = 0;
            row = row + 1;
        }
        parser.Close();

        // Writes the default import file header to row 0
        data = RedoDefaultHeader(data);

        // Writing to file
        WriteCSV("import.csv", data);
    }

    // This function will need additional work. It was created when I was annoyed by Excel regarding a specific file.
    static void CorrectZip()
    {
        Console.WriteLine("Specify the file that will be changed: ");
        string filename = Console.ReadLine();
        int lineCount = File.ReadAllLines(filename).Length;
        Console.WriteLine("Processing " + lineCount + " lines.");
        string[,] data = new string[lineCount, 76];
        TextFieldParser parser = new TextFieldParser(filename);

        parser.HasFieldsEnclosedInQuotes = true;
        parser.SetDelimiters(",");
        string[] fields;
        int row = 0;
        int column = 0;
        while (!parser.EndOfData)
        {
            fields = parser.ReadFields();
            foreach (string field in fields)
            {
                //Console.WriteLine(field); // This can be used to see data as it loads into the data array
                data[row, column] = "\"" + field + "\"";
                if (column == 22 || column == 27)
                {
                    //Console.Write(field + " - ");
                    if (field.Length <= 4 && field.Length != 0) // Add a 0 in front if the field is less than 5 characters, but more than 0 characters
                    {
                        //Console.WriteLine(row + ", " + column + ": " + field);
                        data[row, column] = "0" + field;
                    }
                }
                else if (column == 36)
                {
                    if (field.Length == 0)
                    {
                        data[row, column] = "0"; // This adds a '0' to the TaxIncluded field if the column is empty.
                    }
                }
                column = column + 1;
            }
            column = 0;
            row = row + 1;
        }
        parser.Close();

        // Writes the default import file header to row 0
        data = RedoDefaultHeader(data);

        // Writing to file
        WriteCSV(filename, data);
    }

    static string[,] RedoDefaultHeader(string[,] data)
    {
        data[0, 0] = "ProcessCode"; data[0, 1] = "DocCode"; data[0, 2] = "DocType"; data[0, 3] = "DocDate"; data[0, 4] = "CompanyCode"; data[0, 5] = "CustomerCode"; data[0, 6] = "EntityUseCode"; data[0, 7] = "LineNo"; data[0, 8] = "TaxCode"; data[0, 9] = "TaxDate"; data[0, 10] = "ItemCode"; data[0, 11] = "Description"; data[0, 12] = "Qty"; data[0, 13] = "Amount"; data[0, 14] = "Discount"; data[0, 15] = "Ref1"; data[0, 16] = "Ref2"; data[0, 17] = "ExemptionNo"; data[0, 18] = "RevAcct"; data[0, 19] = "DestAddress"; data[0, 20] = "DestCity"; data[0, 21] = "DestRegion"; data[0, 22] = "DestPostalCode"; data[0, 23] = "DestCountry"; data[0, 24] = "OrigAddress"; data[0, 25] = "OrigCity"; data[0, 26] = "OrigRegion"; data[0, 27] = "OrigPostalCode"; data[0, 28] = "OrigCountry"; data[0, 29] = "LocationCode"; data[0, 30] = "SalesPersonCode"; data[0, 31] = "PurchaseOrderNo"; data[0, 32] = "CurrencyCode"; data[0, 33] = "ExchangeRate"; data[0, 34] = "ExchangeRateEffDate"; data[0, 35] = "PaymentDate"; data[0, 36] = "TaxIncluded"; data[0, 37] = "DestTaxRegion"; data[0, 38] = "OrigTaxRegion"; data[0, 39] = "Taxable"; data[0, 40] = "TaxType"; data[0, 41] = "TotalTax"; data[0, 42] = "CountryName"; data[0, 43] = "CountryCode"; data[0, 44] = "CountryRate"; data[0, 45] = "CountryTax"; data[0, 46] = "StateName"; data[0, 47] = "StateCode"; data[0, 48] = "StateRate"; data[0, 49] = "StateTax"; data[0, 50] = "CountyName"; data[0, 51] = "CountyCode"; data[0, 52] = "CountyRate"; data[0, 53] = "CountyTax"; data[0, 54] = "CityName"; data[0, 55] = "CityCode"; data[0, 56] = "CityRate"; data[0, 57] = "CityTax"; data[0, 58] = "Other1Name"; data[0, 59] = "Other1Code"; data[0, 60] = "Other1Rate"; data[0, 61] = "Other1Tax"; data[0, 62] = "Other2Name"; data[0, 63] = "Other2Code"; data[0, 64] = "Other2Rate"; data[0, 65] = "Other2Tax"; data[0, 66] = "Other3Name"; data[0, 67] = "Other3Code"; data[0, 68] = "Other3Rate"; data[0, 69] = "Other3Tax"; data[0, 70] = "Other4Name"; data[0, 71] = "Other4Code"; data[0, 72] = "Other4Rate"; data[0, 73] = "Other4Tax"; data[0, 74] = "ReferenceCode"; data[0, 75] = "BuyersVATNo";
        return data;
    }

    // Writing to file
    /*
     * 
     * Enhancement: Automatically split the files created into less than 100,000 line files.
     *
    */
    static void WriteCSV(string filename, string[,] data)
    {
        int column = 0;

        System.IO.File.WriteAllText(filename, string.Empty); // Before writing to the file, this empties the file. This way if there were previous contents with more lines than we are writing now, we will not have any of the old contents.
        try
        {
            var fs = File.Open(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            var sw = new StreamWriter(fs);
            column = 0;
            foreach (string field in data)
            {
                if (column <= 74)
                {
                    sw.Write(field + ",");
                    column = column + 1;
                    //break;
                }
                else
                {
                    sw.WriteLine(field);
                    column = 0;
                    //break;
                }
            }
            sw.Flush();
            fs.Close();
        }
        catch (Exception e)
        {
            Console.WriteLine("Exception: " + e.Message);
        }
        Console.WriteLine("Done.");
    }
}

/*
 * ----------------Notes----------------
 * Import Template:
 * ProcessCode,DocCode,DocType,DocDate,CompanyCode,CustomerCode,EntityUseCode,LineNo,TaxCode,TaxDate,ItemCode,Description,Qty,Amount,Discount,Ref1,Ref2,ExemptionNo,RevAcct,DestAddress,DestCity,DestRegion,DestPostalCode,DestCountry,OrigAddress,OrigCity,OrigRegion,OrigPostalCode,OrigCountry,LocationCode,SalesPersonCode,PurchaseOrderNo,CurrencyCode,ExchangeRate,ExchangeRateEffDate,PaymentDate,TaxIncluded,DestTaxRegion,OrigTaxRegion,Taxable,TaxType,TotalTax,CountryName,CountryCode,CountryRate,CountryTax,StateName,StateCode,StateRate,StateTax,CountyName,CountyCode,CountyRate,CountyTax,CityName,CityCode,CityRate,CityTax,Other1Name,Other1Code,Other1Rate,Other1Tax,Other2Name,Other2Code,Other2Rate,Other2Tax,Other3Name,Other3Code,Other3Rate,Other3Tax,Other4Name,Other4Code,Other4Rate,Other4Tax,ReferenceCode,BuyersVATNo
 * 76 fields
*/