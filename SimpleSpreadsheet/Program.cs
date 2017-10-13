using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleSpreadsheet
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter table size:");
            string size = Console.ReadLine();
            try
            {
                Sheet sheet = new Sheet(size);
                string inputString = "";
                Console.WriteLine("Please input cell value:");
                inputString = Console.ReadLine();
                while (inputString != "")
                {
                    string errorMessage;
                    if (!sheet.SetCellValue(inputString, out errorMessage))
                    {
                        Console.WriteLine(errorMessage);
                    }
                    Console.WriteLine("Please input another cell value, or press enter to output result:");
                    inputString = Console.ReadLine();
                }

                Console.WriteLine(sheet.getHeader());

                for (int i = 0; i < sheet.SheetData.GetLength(0); i++)
                {
                    Console.Write(string.Format("{0}\t\t", i + 1));
                    for (int j = 0; j < sheet.SheetData.GetLength(1); j++)
                    {
                        Console.Write(string.Format("{0}\t\t", sheet.SheetData[i, j]));
                    }
                    Console.Write(Environment.NewLine + Environment.NewLine);
                }
            }
            catch (ArgumentException ae) {
                Console.WriteLine(ae.Message);
            }
    
            Console.ReadLine();
        }
    }
}
