using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace AvaTaxImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Shows the openFileDialog
            openFileDialog1.ShowDialog();
            //Reads the text file
            System.IO.StreamReader OpenFile = new System.IO.StreamReader(openFileDialog1.FileName);
            //Displays the text file in the textBox
            textBox1.Text = OpenFile.ReadToEnd();
            //Stores the full path for the file and file name for later use
            FilePath.Text = openFileDialog1.FileName.ToString();
            //Closes the process
            OpenFile.Close();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = "Ava_Import";
            sfd.Filter = "CSV (*.csv)|*.csv";
            //Open the saveFileDialog
            sfd.ShowDialog();
            //Determines the text file to save to
            if (sfd.FileName != null)
            {
                System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(sfd.FileName);
                //Writes the text to the file
                SaveFile.WriteLine(textBox1.Text);
                //Closes the process
                SaveFile.Close();
            }
        }
        //----------------------------------------------------------------
        private string[] FormatSwitch()
        {
            DoesMapExist();
            string[] SwitchTable = new string[76];
            TextFieldParser parser = new TextFieldParser("mapping.ini");
            parser.SetDelimiters(",");
            string[] fields;
            int row = 0;
            while (!parser.EndOfData)
            {
                fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    SwitchTable[row] = field;
                    if(field == null)
                    { MessageBox.Show(field); }
                }
                row = row + 1;
            }
            return SwitchTable;
        }

        private void DoesMapExist()
        {
            if (!File.Exists("mapping.ini"))
            {
                // This is a hard-coded version of mapping to go from the Export Document Lines report format to the import format.
                string[] SwitchTable = new string[76];
                SwitchTable[0] = "0"; // Process Code
                SwitchTable[1] = "4"; // Company Code
                SwitchTable[2] = "0";
                SwitchTable[3] = "1"; // DocCode
                SwitchTable[4] = "2"; // DocType
                SwitchTable[5] = "3"; // ?
                SwitchTable[6] = "9"; // TaxDate
                SwitchTable[7] = "0"; // N/A
                SwitchTable[8] = "0"; // N/A
                SwitchTable[9] = "0"; // N/A
                SwitchTable[10] = "0"; // N/A
                SwitchTable[11] = "32"; // ?
                SwitchTable[12] = "5"; // ?
                SwitchTable[13] = "6"; // ?
                SwitchTable[14] = "17"; // ?
                SwitchTable[15] = "31"; // ?
                SwitchTable[16] = "0"; // N/A
                SwitchTable[17] = "30"; // ?
                SwitchTable[18] = "29"; // ?
                SwitchTable[19] = "7"; // ?
                SwitchTable[20] = "10"; // ?
                SwitchTable[21] = "11"; // ?
                SwitchTable[22] = "8"; // ?
                SwitchTable[23] = "0"; // N/A
                SwitchTable[24] = "0"; // N/A
                SwitchTable[25] = "12"; // ?
                SwitchTable[26] = "13"; // Line Amount
                SwitchTable[27] = "14"; // ?
                SwitchTable[28] = "15"; // ?
                SwitchTable[29] = "16"; // ?
                SwitchTable[30] = "18"; // ?
                SwitchTable[31] = "0"; // N/A
                SwitchTable[32] = "0"; // N/A
                SwitchTable[33] = "0"; // N/A
                SwitchTable[34] = "41"; // Tax Amount & Tax Includes
                SwitchTable[35] = "0"; // N/A
                SwitchTable[36] = "0"; // N/A
                SwitchTable[37] = "0"; // N/A
                SwitchTable[38] = "0"; // N/A
                SwitchTable[39] = "24"; // ?
                SwitchTable[40] = "25"; // ?
                SwitchTable[41] = "26"; // ?
                SwitchTable[42] = "28"; // ?
                SwitchTable[43] = "27"; // ?
                SwitchTable[44] = "0"; // N/A
                SwitchTable[45] = "19"; // ?
                SwitchTable[46] = "20"; // ?
                SwitchTable[47] = "21"; // ?
                SwitchTable[48] = "23"; // ?
                SwitchTable[49] = "22"; // ?
                Write_ini("mapping.ini", SwitchTable);
            }
        }

        private string[] ReadForImportant(string[] mapping) // This function is used to detect the fields that must be handled separately later.
        {
            string[] HandleSep = new string[6];
            int row = 0;
            while (row < 76)
            {
                if (mapping[row] == "0") // Detects for the Process Code
                { HandleSep[0] = row.ToString(); }
                else if (mapping[row] == "2") // Detects for the DocType
                { HandleSep[1] = row.ToString(); }
                else if (mapping[row] == "4") // Detects for the Company Code
                { HandleSep[2] = row.ToString(); }
                else if (mapping[row] == "9") // Detects for the Tax Date, needed for creating inverse
                { HandleSep[3] = row.ToString(); }
                else if (mapping[row] == "13") // Detects for Line Amount
                { HandleSep[4] = row.ToString(); }
                else if (mapping[row] == "41") // Detects for Tax Amount
                { HandleSep[5] = row.ToString(); }
                row = row + 1;
            }
            return HandleSep;
        }

        private void ConvertExport(string ProcessCode, string CompanyCode, bool taxincludes)
        {
            //int lineCount = textBox1.Lines.Count();
            //lineCount -= String.IsNullOrWhiteSpace(textBox1.Text) ? 1 : 0;
            int lineCount = File.ReadAllLines(FilePath.Text).Length;
            //Console.WriteLine("Processing " + lineCount + " lines."); - This needs to be replaced with the Windows Form version
            string[,] data = new string[lineCount, 76];
            TextFieldParser parser = new TextFieldParser(FilePath.Text); // File name required in quotes

            parser.HasFieldsEnclosedInQuotes = true;
            parser.SetDelimiters(",");
            string[] fields;
            int row = 0;
            int column = 0;
            bool taxdateoverride = false;
            double line_amount;
            double tax_amount;
            string ti_string;
            string[] ST = new string[76];
            ST = FormatSwitch();
            string[] HandleSep = new string[6];
            HandleSep = ReadForImportant(ST);
            while (!parser.EndOfData)
            {
                fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    data[0, 0] = field; // testing
                    if (Convert.ToInt32(ST[column]) != 0) // We must skip over column 0 and the undifined columns
                    { data[row, Convert.ToInt32(ST[column])] = "\"" + field + "\""; }// This uses the previously defined mapping to reorgainize the data in the "data" array.
                    if(column == Convert.ToInt32(HandleSep[0]))
                    {
                        data[row, 0] = ProcessCode;
                    }
                    else if(column == Convert.ToInt32(HandleSep[1]))
                    {
                        //data[row, 2] = "0"; //This is a temporary change, which should be rolled into an enhancement
                        if (field == "Sales Invoice")
                        { data[row, 2] = "1"; }
                        else if (field == "Return Invoice")
                        {
                            data[row, 2] = "5";
                            taxdateoverride = true;
                        }
                        else { data[row, 2] = "1"; } // This assumes that if the data isn't "Sales Invoice" or "Return Invoice", that the intention was for this to be an invoice
                    }
                    else if(column == Convert.ToInt32(HandleSep[2]))
                    {
                        data[row, 4] = "\"" + CompanyCode + "\"";
                    }
                    else if(column == Convert.ToInt32(HandleSep[3]))
                    {
                        if (taxdateoverride)
                            {
                                data[row, 9] = field; // This will set the date override date to the date used in the original, but only for return invoices.
                                taxdateoverride = false;
                            }
                    }
                    else if(column == Convert.ToInt32(HandleSep[5]))
                    {
                        if (taxincludes && row != 0) // This looks for the tax includes flag, which can be set by the app user, and it skips the header row.
                            {
                                tax_amount = Convert.ToDouble(field);
                                line_amount = Convert.ToDouble(data[row, 13]);
                                line_amount = line_amount + tax_amount;
                                ti_string = Convert.ToString(line_amount);
                                data[row, 13] = ti_string; // This makes the line amount equal to the original amount plus the tax
                                data[row, 36] = "1"; // This sets the Tax Includes field to True on the import
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
            textBox1.Text = UpdateFilePreview(data);
        }

        private string[,] RedoDefaultHeader(string[,] data)
        {
            data[0, 0] = "ProcessCode"; data[0, 1] = "DocCode"; data[0, 2] = "DocType"; data[0, 3] = "DocDate"; data[0, 4] = "CompanyCode"; data[0, 5] = "CustomerCode"; data[0, 6] = "EntityUseCode"; data[0, 7] = "LineNo"; data[0, 8] = "TaxCode"; data[0, 9] = "TaxDate"; data[0, 10] = "ItemCode"; data[0, 11] = "Description"; data[0, 12] = "Qty"; data[0, 13] = "Amount"; data[0, 14] = "Discount"; data[0, 15] = "Ref1"; data[0, 16] = "Ref2"; data[0, 17] = "ExemptionNo"; data[0, 18] = "RevAcct"; data[0, 19] = "DestAddress"; data[0, 20] = "DestCity"; data[0, 21] = "DestRegion"; data[0, 22] = "DestPostalCode"; data[0, 23] = "DestCountry"; data[0, 24] = "OrigAddress"; data[0, 25] = "OrigCity"; data[0, 26] = "OrigRegion"; data[0, 27] = "OrigPostalCode"; data[0, 28] = "OrigCountry"; data[0, 29] = "LocationCode"; data[0, 30] = "SalesPersonCode"; data[0, 31] = "PurchaseOrderNo"; data[0, 32] = "CurrencyCode"; data[0, 33] = "ExchangeRate"; data[0, 34] = "ExchangeRateEffDate"; data[0, 35] = "PaymentDate"; data[0, 36] = "TaxIncluded"; data[0, 37] = "DestTaxRegion"; data[0, 38] = "OrigTaxRegion"; data[0, 39] = "Taxable"; data[0, 40] = "TaxType"; data[0, 41] = "TotalTax"; data[0, 42] = "CountryName"; data[0, 43] = "CountryCode"; data[0, 44] = "CountryRate"; data[0, 45] = "CountryTax"; data[0, 46] = "StateName"; data[0, 47] = "StateCode"; data[0, 48] = "StateRate"; data[0, 49] = "StateTax"; data[0, 50] = "CountyName"; data[0, 51] = "CountyCode"; data[0, 52] = "CountyRate"; data[0, 53] = "CountyTax"; data[0, 54] = "CityName"; data[0, 55] = "CityCode"; data[0, 56] = "CityRate"; data[0, 57] = "CityTax"; data[0, 58] = "Other1Name"; data[0, 59] = "Other1Code"; data[0, 60] = "Other1Rate"; data[0, 61] = "Other1Tax"; data[0, 62] = "Other2Name"; data[0, 63] = "Other2Code"; data[0, 64] = "Other2Rate"; data[0, 65] = "Other2Tax"; data[0, 66] = "Other3Name"; data[0, 67] = "Other3Code"; data[0, 68] = "Other3Rate"; data[0, 69] = "Other3Tax"; data[0, 70] = "Other4Name"; data[0, 71] = "Other4Code"; data[0, 72] = "Other4Rate"; data[0, 73] = "Other4Tax"; data[0, 74] = "ReferenceCode"; data[0, 75] = "BuyersVATNo";
            return data;
        }

        private string UpdateFilePreview(string[,] data)
        {
            int column = 0;
            string ConvertedOutput = "";
            column = 0;
            foreach (string field in data)
            {
                if (column <= 74)
                {
                    ConvertedOutput = ConvertedOutput + field + ",";
                    column = column + 1;
                    //break;
                }
                else
                {
                    ConvertedOutput = ConvertedOutput + "\n";
                    column = 0;
                    //break;
                }
            }
            return ConvertedOutput;
        }

        static void Write_ini(string filename, string[] data)
        {
            System.IO.File.WriteAllText(filename, string.Empty); // Before writing to the file, this empties the file. This way if there were previous contents with more lines than we are writing now, we will not have any of the old contents.
            try
            {
                var fs = File.Open(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                var sw = new StreamWriter(fs);
                foreach (string field in data)
                {
                    sw.WriteLine(field);
                }
                sw.Flush();
                fs.Close();
            }
            catch (Exception e)
            {
                // Console.WriteLine("Exception: " + e.Message); // Replace with Windows Form messaging
            }
            // Console.WriteLine("Done.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConvertExport("4", "", false);
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }
        //----------------------------------------------------------------
    }
}
