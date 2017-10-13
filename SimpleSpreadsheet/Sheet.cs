using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SimpleSpreadsheet
{
    public class Sheet
    {
        private class CellReference
        {
            public int RowIndex { get; set; }
            public int ColIndex { get; set; }
            public string Reference { get; set; }
        }

        private string[,] table;
        Dictionary<int, string> ColNames = new Dictionary<int, string>();
        List<CellReference> CellReferences = new List<CellReference>();
        private int colNum;

        public string[,] SheetData
        {
            get { return table; }
        }

        public Sheet(string size)
        {
            string[] splittedInput = size.Split('*');
            if (splittedInput.Length == 2)
            {
                int rows;
                int cols;
                if (int.TryParse(splittedInput[0].Trim(), out rows) && int.TryParse(splittedInput[0].Trim(), out cols))
                {
                    // Size in correct format, initialise the empty table provided dimension
                    table = new string[rows, cols];
                    colNum = cols;
                    for (int i = 0; i < cols; i++)
                    {
                        for (int j = 0; j < rows; j++)
                        {
                            CellReferences.Add(new CellReference
                            {
                                RowIndex = j,
                                ColIndex = i,
                                Reference = GetColName(i) + (j + 1)
                            });
                        }
                    }
                }
                else {
                    throw new ArgumentException("Invalid table size.");
                }
            }
            else {
                throw new ArgumentException("Invalid table size.");
            }
        }

        /// <summary>
        /// Show column header as A B C etc
        /// </summary>
        private string GetColName(int colIndex)
        {
            int dividend = colIndex + 1;
            string colName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                colName = Convert.ToChar(65 + modulo).ToString() + colName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return colName;
        }

        /// <summary>
        /// Output header row to screen
        /// </summary>
        public string getHeader()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\t\t");
            for (int j = 0; j < colNum; j++)
            {
                sb.AppendFormat("{0}\t\t", GetColName(j));
            }

            return sb.ToString();
        }

        /// <summary>
        /// evaluate input string to work out correct cell and value to populate
        /// </summary>
        public bool SetCellValue(string input, out string errorMessage)
        {
            bool retVal = false;
            errorMessage = "";
            // Parse input string to cell reference and input vaule
            string[] splittedInput = input.Split('=');
            if (splittedInput.Length == 2)
            {
                CellReference cell = LookupCell(splittedInput[0].Trim());
                if (cell == null)
                {
                    errorMessage = string.Format("Cell {0} does not exist", splittedInput[1]);
                }
                else
                {
                    //resolve value to a number, either from a string number or from cell reference
                    string value = splittedInput[1].Trim();
                    int numberValue;
                    if (!int.TryParse(value, out numberValue))
                    {
                        if (EvaluateInput(value, out numberValue))
                        {
                            table[cell.RowIndex, cell.ColIndex] = numberValue.ToString();
                            retVal = true;
                        }
                        else
                        {
                            errorMessage = "Invalid input";
                        }
                    }
                    else {
                        table[cell.RowIndex, cell.ColIndex] = numberValue.ToString();
                        retVal = true;
                    }
                }
            }
            else {
                errorMessage = "Invalid input";
            }
            
            return retVal;
        }

        /// <summary>
        /// Identify cell details based on reference string, ie A1, A2, B1, B2
        /// </summary>
        private CellReference LookupCell(string reference) {
            return CellReferences.FirstOrDefault(c => c.Reference.Equals(reference, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// resolve value to a number, either from a string number or from cell reference
        /// </summary>
        private bool EvaluateInput(string input, out int output) {
            bool retVal = false;
            output = 0;
            char[] operators =  new char[]{ '+', '-', '*', '/'};
            for (int i = 0; i < operators.Length; i++) {
                string[] splittedInput = input.Split(operators[i]);
                if (splittedInput.Length == 2)
                {
                    int leftValue;
                    if (!int.TryParse(splittedInput[0].Trim(), out leftValue))
                    {
                        CellReference cell = LookupCell(splittedInput[0].Trim());
                        if (cell != null)
                        {
                            leftValue = int.Parse(table[cell.RowIndex, cell.ColIndex]);
                        }
                        else
                        {
                            retVal = false;
                            break;
                        }
                    }

                    int rightValue;
                    if (!int.TryParse(splittedInput[1].Trim(), out rightValue))
                    {
                        CellReference cell = LookupCell(splittedInput[1].Trim());
                        if (cell != null)
                        {
                            rightValue = int.Parse(table[cell.RowIndex, cell.ColIndex]);
                        }
                        else
                        {
                            retVal = false;
                            break;
                        }
                    }

                    string mathExp = leftValue.ToString() + operators[i].ToString() + rightValue.ToString();

                    try
                    {
                        object result = new DataTable().Compute(mathExp, null);
                        output = Convert.ToInt32(result);
                        retVal = true;
                        break;
                    }
                    catch (Exception ex) {
                        retVal = false;
                        break;
                    }
                }
            }

            return retVal;
        }
    }
}
