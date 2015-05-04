/*
 * For Excel 2007 or 2010 or later (maybe - not tested)
  * Install 2007/2010 Office System Driver - http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
 * 
 */

using System.Text;
using Parser;

namespace ParserTest
{
    /// <summary>
    /// These Simple programs are for testing the Parser classes
    /// Uncomment the program that matched the functionality you want to test and complile
    /// </summary>
    class Program
    {

        /// <summary>
        /// Parse Excel to CSV File
        /// </summary>
        /// <param name="args"></param>
        //static void Main(string[] args)
        //{
        //    StringBuilder outputText = new StringBuilder();
        //    var parser = new Parse();
        //    CsvDto csvDto = parser.ParseExcelFile((args[0]));
        //    outputText.Append(csvDto.CsvData);

        //    using (StreamWriter writer = new StreamWriter(args[1]))
        //        writer.Write(outputText);
        //}



        /// <summary>
        /// Parse CSV to Excel file
        /// </summary>
        /// <param name="args">args[0] id the UNC path of the CSV input file, args[1] is the UNC path of the output .xlsx file</param>
        static void Main(string[] args)
        {
            StringBuilder outputText = new StringBuilder();
            
            var csvParser = new Parser.CSVParser(true, 9999);
            var csvDto = csvParser.ParseCsvFile(args[0], ",", "'", true,true);
            var parser = new ExcelParser();
            parser.WriteExcelFileFromCSV(args[1], csvDto.CsvData, true);
            
        }
    }
}
