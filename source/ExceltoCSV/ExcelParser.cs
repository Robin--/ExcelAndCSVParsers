/*
 * For Excel 2007 or 2010 or later (maybe - not tested)
  * Install 2007/2010 Office System Driver - http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
 * 
 * 
 * 
 * Todo Add extra paramter to handle incluuded record header
 * Extended Properties=""Text;HDR=No;FMT=Delimited\"""
 */

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace Parser
{
    public class ExcelParser
    {
        private static String[] GETWorksheetList(String connectionString)
        {
            OleDbConnection objConn = null;
            DataTable sheets = null;
            try
            {
                objConn = new OleDbConnection(connectionString);
                objConn.Open(); // Open connection with the database.

                // Get the data table containing the schema guid.
                sheets = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Add the sheet name to the string array.
                int k = 0;
                String temp;
                if (sheets != null)
                {
                    String[] worksheets = new String[sheets.Rows.Count];
                    foreach (DataRow row in sheets.Rows)
                    {
                        temp = row["TABLE_NAME"].ToString();
                        temp = temp.Replace("'", "");
                        worksheets[k] = temp.Substring(0, temp.Length - 1);
                        k++;
                    }
                    return worksheets;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                // Clean up.

                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (sheets != null)
                    sheets.Dispose();
            }

            return null;
        }

        private StringBuilder returnCSVRecord(string connectionString, String worksheetName)
        {
            StringBuilder outputText = new StringBuilder();
            try
            {
                //Fill the dataset with information from the Sheet 1 worksheet.
                var adapter1 = new OleDbDataAdapter("SELECT * FROM [" + worksheetName + "$]", connectionString);
                var ds = new DataSet();
                adapter1.Fill(ds, "results");
                DataTable data = ds.Tables["results"];

                //Show all columns
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    string record = string.Empty;
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        if (j + 1 == data.Columns.Count)
                            record += "\"" + data.Rows[i].ItemArray[j] + "\"";
                        else
                            record += "\"" + data.Rows[i].ItemArray[j] + "\",";
                    }

                    outputText.AppendLine(record);
                }
            }
            catch (Exception ex)
            {
                outputText.AppendLine("Error with Record");
            }

            return outputText;
        }

        public CsvDto ParseExcelFileToCSV(string ExcelFileName)
        {
            var csvDto = new CsvDto();

            try
            {
                StringBuilder outputText = new StringBuilder();
                //Create a connection string to access the Excel file using the ACE provider.
                //This is for Excel 2007 + 2010. 2003 uses an older driver.
                //var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;IMEX=1;HDR=NO;", ExcelFileName);
                var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=NO;IMEX=1'", ExcelFileName);
                // OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + PrmPathExcelFile + @";Extended Properties=""Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text""");

                // Read the Worksheets
                // This codebase depends on the columns being the same on each sheet
                String[] worksheetList = GETWorksheetList(connectionString);
                if (worksheetList != null)
                    foreach (String worksheetName in worksheetList)
                    {
                        outputText.Append(returnCSVRecord(connectionString, worksheetName));
                    }

                csvDto.CsvData = outputText.ToString();
            }
            catch (Exception)
            {
                csvDto.CsvData = "File format error Could not parse the file to Excel";
            }


            return csvDto;
        }


        /// <summary>
        /// Takes in a CSV file in a string, and creates a Excel File from it
        /// </summary>
        /// <param name="ExcelFileName"> The UNC path for the Excel file to write</param>
        /// <param name="CSVData">the string containing the CSV data to parse</param>
        /// <param name="overWrite">Overwrite the Excel file?</param>
        /// <returns></returns>
        public CsvDto WriteExcelFileFromCSV(string ExcelFileName, string CSVData, bool overWrite)
        {
            var csvDto = new CsvDto();

            try
            {
                var csvData = StructureCsvContent(CSVData);
                var olecon = new OleDbConnection();
                var olecommand = new OleDbCommand();

                if (overWrite)
                {
                    try
                    {
                        if (File.Exists(ExcelFileName))
                            File.Delete(ExcelFileName);
                    }
                    catch (Exception)
                    {
                        csvDto.CsvData = "The Excel file is currently open with another program. Please close the file before running";
                        return csvDto;
                    }

                }

                var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES'", ExcelFileName);

                olecon.ConnectionString = connectionString;

                try
                {
                    olecon.Open();
                    olecommand.Connection = olecon;
                }
                catch (Exception)
                {
                    csvDto.CsvData = "This parser only supports Excel files from 2007 onwards.";
                    return csvDto;
                }
                olecon.Open();
                olecommand.Connection = olecon;


                string commandtext = string.Empty;
                string querytext = string.Empty;
                for (int i = 0; i < csvData[0].Length; i++)
                {
                    if (i == (csvData[0].Length - 1))
                    {
                        commandtext += csvData[0][i] + " VARCHAR";
                        querytext += csvData[0][i];
                    }
                    else
                    {
                        commandtext += csvData[0][i] + " VARCHAR, ";
                        querytext += csvData[0][i] + ", ";
                    }

                }

                // Create the Worksheet
                var buildquery = "CREATE TABLE Sheet1 (" + commandtext + ")";
                olecommand.CommandText = buildquery;
                olecommand.ExecuteNonQuery();

                //example
                //olecommand.CommandText = "CREATE TABLE Sheet1 (Sno Int, Employee_Name VARCHAR, Company VARCHAR, Date_Of_joining DATE, Stipend DECIMAL, Stocks_Held DECIMAL)";


                for (int i = 1; i < csvData.Count; i++)
                {
                    string insertrow = "INSERT INTO Sheet1 (" + querytext + ") values ('";

                    for (int j = 0; j < csvData[0].Length; j++)
                    {
                        if (j == (csvData[0].Length - 1))
                            insertrow += csvData[i][j] + "'";
                        else
                            insertrow += csvData[i][j] + "','";
                    }
                    insertrow += ")";
                    olecommand.CommandText = insertrow;
                    olecommand.ExecuteNonQuery();
                }

                //example
                //olecommand.CommandText = "INSERT INTO Sheet1 (Sno, Employee_Name, Company,Date_Of_joining,Stipend,Stocks_Held) values ('1', 'Siddharth Rout', 'Defining Horizons', '20/7/2014','2000.75','0.01')";

                olecon.Close();


                csvDto.CsvData = "Pass";
            }
            catch (Exception)
            {
                csvDto.CsvData = "Failed to parse the file due to malformed data";
            }

            return csvDto;

        }

        /// <summary>
        /// Parses a CSV recordset contained in a string into a list containing the CSV fiields split into a string array
        /// </summary>
        /// <param name="csvData"></param>
        /// <returns></returns>
        private List<string[]> StructureCsvContent(string csvData)
        {
            byte[] arr = Encoding.ASCII.GetBytes(csvData);
            var lines = new List<string[]>();

            using (MemoryStream memStream = new MemoryStream(arr, false))
            using (StreamReader lReader = new StreamReader(memStream))
                while (!lReader.EndOfStream)
                {
                    var readLine = lReader.ReadLine();
                    if (readLine != null)
                    {
                        var line = readLine.Split(',');
                        var i = 0;
                        foreach (var field in line)
                        {
                            var cleanstring = string.Empty;

                            if (field != null)
                                foreach (var lLetter in field)
                                    if (lLetter != '"')
                                        cleanstring += lLetter;

                            line[i] = cleanstring;
                            i++;
                        }

                        lines.Add(line);
                    }
                }
            return lines;
        }
    }


    /// <summary>
    /// Return class that ontains a single string
    /// </summary>
    public class CsvDto
    {
        public string CsvData { get; internal set; }
    }
}