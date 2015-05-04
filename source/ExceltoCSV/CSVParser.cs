using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Parser
{
    public class CSVParser
    {

        static char _commaCharacter = ',';
        static char _quoteCharacter = '"';
        
        private bool TrimTrailingEmptyLines { get; set; }
        private int MaxColumnsToRead { get; set; }

        public CSVParser(bool trimTrailingEmptyLines, int maxColumnsToRead)
        {
            MaxColumnsToRead = maxColumnsToRead;
            TrimTrailingEmptyLines = trimTrailingEmptyLines;
        }

        #region Nested types

        private abstract class ParserState
        {
            public static readonly LineStartState LineStartState = new LineStartState();
            public static readonly ValueStartState ValueStartState = new ValueStartState();
            public static readonly ValueState ValueState = new ValueState();
            public static readonly QuotedValueState QuotedValueState = new QuotedValueState();
            public static readonly QuoteState QuoteState = new QuoteState();

            public abstract ParserState AnyChar(char ch, ParserContext context);
            public abstract ParserState Comma(ParserContext context);
            public abstract ParserState Quote(ParserContext context);
            public abstract ParserState EndOfLine(ParserContext context);
        }

        private class LineStartState : ParserState
        {
            public override ParserState AnyChar(char ch, ParserContext context)
            {
                context.AddChar(ch);
                return ValueState;
            }

            public override ParserState Comma(ParserContext context)
            {
                context.AddValue();
                return ValueStartState;
            }

            public override ParserState Quote(ParserContext context)
            {
                return QuotedValueState;
            }

            public override ParserState EndOfLine(ParserContext context)
            {
                context.AddLine();
                return LineStartState;
            }
        }

        private class ValueStartState : LineStartState
        {
            public override ParserState EndOfLine(ParserContext context)
            {
                context.AddValue();
                context.AddLine();
                return LineStartState;
            }
        }

        private class ValueState : ParserState
        {
            public override ParserState AnyChar(char ch, ParserContext context)
            {
                context.AddChar(ch);
                return ValueState;
            }

            public override ParserState Comma(ParserContext context)
            {
                context.AddValue();
                return ValueStartState;
            }

            public override ParserState Quote(ParserContext context)
            {
                context.AddChar(_quoteCharacter);
                return ValueState;
            }

            public override ParserState EndOfLine(ParserContext context)
            {
                context.AddValue();
                context.AddLine();
                return LineStartState;
            }
        }

        private class QuotedValueState : ParserState
        {
            public override ParserState AnyChar(char ch, ParserContext context)
            {
                context.AddChar(ch);
                return QuotedValueState;
            }

            public override ParserState Comma(ParserContext context)
            {
                context.AddChar(_commaCharacter);
                return QuotedValueState;
            }

            public override ParserState Quote(ParserContext context)
            {
                return QuoteState;
            }

            public override ParserState EndOfLine(ParserContext context)
            {
                context.AddChar('\r');
                context.AddChar('\n');
                return QuotedValueState;
            }
        }

        private class QuoteState : ParserState
        {
            public override ParserState AnyChar(char ch, ParserContext context)
            {
                //undefined, ignore "
                context.AddChar(ch);
                return QuotedValueState;
            }

            public override ParserState Comma(ParserContext context)
            {
                context.AddValue();
                return ValueStartState;
            }

            public override ParserState Quote(ParserContext context)
            {
                context.AddChar(_quoteCharacter);
                return QuotedValueState;
            }

            public override ParserState EndOfLine(ParserContext context)
            {
                context.AddValue();
                context.AddLine();
                return LineStartState;
            }
        }

        private class ParserContext
        {
            private readonly StringBuilder _currentValue = new StringBuilder();
            private readonly List<string[]> _lines = new List<string[]>();
            private readonly List<string> _currentLine = new List<string>();

            public ParserContext()
            {
                MaxColumnsToRead = 1000;
            }

            public int MaxColumnsToRead { get; set; }

            public void AddChar(char ch)
            {
                _currentValue.Append(ch);
            }

            public void AddValue()
            {
                if (_currentLine.Count < MaxColumnsToRead)
                    _currentLine.Add(_currentValue.ToString());
                _currentValue.Remove(0, _currentValue.Length);
            }

            public void AddLine()
            {
                _lines.Add(_currentLine.ToArray());
                _currentLine.Clear();
            }

            public List<string[]> GetAllLines()
            {
                if (_currentValue.Length > 0)
                {
                    AddValue();
                }
                if (_currentLine.Count > 0)
                {
                    AddLine();
                }
                return _lines;
            }
        }

        #endregion



        //public string Parse(string inputFilename, string splitCharacter, string quoteCharacter, bool trimWhiteSpaces, bool useFields)
        public CsvDto ParseCsvFile(string inputFilename, string splitCharacter, string quoteCharacter, bool trimWhiteSpaces, bool useFields)
        {
            var csvDto = new CsvDto();

            _commaCharacter = splitCharacter[0];
            _quoteCharacter = quoteCharacter[0];

            try
            {
                TextReader reader = File.OpenText(inputFilename);
                csvDto = ParseCsv(reader, trimWhiteSpaces, useFields);
                return csvDto;
            }
            catch(Exception)
            {
                csvDto.CsvData = "Could not open the filename provided, or the file does not exist";
                return csvDto;

            }

        }

        private CsvDto ParseCsv(TextReader reader, bool trimWhiteSpaces, bool useFields)
        {

            var csvDto = new CsvDto();

            try
            {
                var context = new ParserContext();
                if (MaxColumnsToRead != 0)
                {
                    context.MaxColumnsToRead = MaxColumnsToRead;
                }

                ParserState currentState = ParserState.LineStartState;
                string next;
                while ((next = reader.ReadLine()) != null)
                {
                    foreach (var ch in next)
                    {
                        if (ch == _commaCharacter)
                        {
                            currentState = currentState.Comma(context);
                        }
                        else if (ch == _quoteCharacter)
                        {
                            currentState = currentState.Quote(context);
                        }
                        else
                        {
                            currentState = currentState.AnyChar(ch, context);
                        }
                    }
                    currentState = currentState.EndOfLine(context);
                }
                List<string[]> allLines = context.GetAllLines();
                if (TrimTrailingEmptyLines && allLines.Count > 0)
                {
                    var isEmpty = true;
                    for (int i = allLines.Count - 1; i >= 0; i--)
                    {
                        for (int j = 0; j < allLines[i].Length; j++)
                        {
                            if (!String.IsNullOrEmpty(allLines[i][j]))
                            {
                                isEmpty = false;
                                break;
                            }
                        }
                        if (!isEmpty)
                        {
                            if(i < allLines.Count - 1)
                                allLines.RemoveRange(i + 1, allLines.Count - i - 1);
                            break;
                        }
                    }
                    if(isEmpty)
                        allLines.RemoveRange(0, allLines.Count);
                }

                StringBuilder outputText = new StringBuilder();

                int linecount = 0;
                foreach (var line in allLines)
                {
                    string ss = string.Empty;
                    foreach (var s in line)
                    {
                        // Remove Unix Control characters from the field values
                        string workingString = new string(s.Where(c => !char.IsControl(c)).ToArray());

                        // Remove any leading and trailing white spaces
                        if(trimWhiteSpaces)
                            workingString = workingString.Trim();

                        if(useFields && linecount == 0)
                            ss += workingString + ",";
                        else
                            ss += "\"" + workingString + "\",";
                    }

                    if(ss.EndsWith(","))
                        ss = ss.Remove(ss.LastIndexOf(",", System.StringComparison.Ordinal));

                    if(useFields && linecount == 0)
                        outputText.AppendLine(ss);
                    else if(linecount > 0)
                        outputText.AppendLine(ss);

                    linecount++;
                }

                csvDto.CsvData = outputText.ToString();
                return csvDto;
            }
            catch(Exception)
            {
                csvDto.CsvData = "The XML was Malformed. Could not parse the file";
                return csvDto;
            }

        }

    }

}