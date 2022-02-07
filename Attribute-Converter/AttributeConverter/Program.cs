using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Config config = Configuration.CreateConfiguration();

            if (config == null)
            {
                Console.WriteLine();
                Console.WriteLine("Konfigurationsdatei 'settings.ini' fehlt.");
                Console.WriteLine();
            }
            else if (config.excelFile == null)
            {
                Console.WriteLine();
                Console.WriteLine("In der Konfigurationsdatei 'settings.ini' fehlt der Parameter 'excelfile'.");
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine("Ausgabesprache: " + config.language.ToString());

                ExcelParser excelParser = new ExcelParser(config);
                IProcessor processor = Factory.CreateInstance(config);
                processor.WriteDocument(excelParser.ProcessExcelFile());

                Console.WriteLine();
            }
        }
    }
}
