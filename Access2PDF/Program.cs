using Access2PDF.Commons;
using Access2PDF.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Access2PDF
{
    class Program
    {
      
        static void Main(string[] args)
        {
            if (Command.ShowHelpCommands(args))
                return;

            var commands = Command.GetModel(args);

            try
            {
                PdfConvert.GenerateFile(commands.ReportName, commands.MsAccess, commands.OutputPdf, commands.Filter);

                Console.ForegroundColor = ConsoleColor.Yellow;

                Console.WriteLine("--------------");
                Console.WriteLine("Arquivo PDF ({0}) exportado com sucesso!", commands.OutputPdf);
                Console.WriteLine("--------------");
            }
            catch(ArgumentException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
            }
            catch(Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;

                Console.WriteLine("Erro ao exportar arquivo ({0} - {1})...{2}", commands.MsAccess, commands.ReportName, ex.Message);

                if (commands.DebugMode)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine(ex);
                }
            }

            Console.ForegroundColor = ConsoleColor.Gray;
        }
    }
}
