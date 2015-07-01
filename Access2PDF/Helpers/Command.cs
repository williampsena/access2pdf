using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Access2PDF.Helpers
{
    public class CommandModel
    {
        /// <summary>
        /// Modo debug
        /// </summary>
        public bool DebugMode { get; set; }

        /// <summary>
        /// Nome do relatório
        /// </summary>
        public string ReportName { get; set;}
        
        /// <summary>
        /// Arquivo de destino PDF
        /// </summary>
        public string OutputPdf { get; set;}

        /// <summary>
        /// Arquivo origem Microsoft Access (*.mdb)
        /// </summary>
        public string MsAccess { get; set;}

        /// <summary>
        /// Filtro do Microsoft Access
        /// </summary>
        public string Filter { get; set; }

        /// <summary>
        /// Lista de todos os argumentos
        /// </summary>
        public Dictionary<string, string> Arguments { get; set;}

        /// <summary>
        /// Construtor padrão
        /// </summary>
        public CommandModel()
        {
            Arguments = new Dictionary<string,string>();
        }
    }

    /// <summary>
    /// Auxilia operações realizadas no console
    /// </summary>
    internal static class Command
    {
        /// <summary>
        /// Expressão regular para captura de comandos do console
        /// </summary>
        private const string REGEXCMD = "^(/|-)(?<name>\\w+)(?:\\:(?<value>.+)$|\\:$|$)";

        /// <summary>
        /// Expressão regular para capturar comando que alterar a cor da fonte do console
        /// </summary>
        private const string REGEXCONSOLECOLOR = "{{color:(?<color>\\w+)}}";

        /// <summary>
        /// Configura a cor da fonte do console
        /// </summary>
        /// <param name="data">Conteúdo</param>
        /// <returns></returns>
        private static bool SetConsoleFontColor(string data)
        {
            var output = false;
            var matchColor = Regex.Match(data, REGEXCONSOLECOLOR, RegexOptions.IgnoreCase);

            if (matchColor.Success)
            {
                var fontColor = ConsoleColor.White;

                if (Enum.TryParse<ConsoleColor>(matchColor.Groups["color"].Value, true, out fontColor))
                {
                    Console.ForegroundColor = fontColor;
                }

                output = true;
            }

            return output;
        }

        /// <summary>
        /// Exibe o conteúdo para ajuda dos comandos do console
        /// </summary>
        /// <param name="args">Argumentos</param>
        /// <returns>Indica se o comando inserido é o comando de ajuda</returns>
        public static bool ShowHelpCommands(string[] args)
        {
            var isHelp = false;
            var firstArgument = args.FirstOrDefault()?? string.Empty;

            if(firstArgument.ToLower().Contains("/help") || firstArgument.StartsWith("?"))
            {
                var consoleFontColor = ConsoleColor.Gray;
                Console.ForegroundColor = consoleFontColor;

                using(var reader = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("Access2PDF.Data.Help.txt")))
                {
                    string line = null;

                    while(!reader.EndOfStream)
                    {
                        line = reader.ReadLine();

                        if (!SetConsoleFontColor(line))
                        {
                            Console.WriteLine(line);
                        }
                    }
                }

                isHelp = true;
            }

            Console.ForegroundColor = ConsoleColor.Gray;

            return isHelp;
        }

        /// <summary>
        /// Obtém o modelo de comandos
        /// </summary>
        /// <param name="args">Argumentos</param>
        /// <returns>Modelo de comando</returns>
        public static CommandModel GetModel(string[] args)
        {
            var model = new CommandModel();

            args.ToList().ForEach(x =>
            {
                var matchCmd = Regex.Match(x.Replace("\"", string.Empty), REGEXCMD, RegexOptions.IgnoreCase);

                if (matchCmd.Success)
                {
                    var name = matchCmd.Groups["name"].Value.ToUpper();
                    var value = matchCmd.Groups["value"].Value;

                    switch (name)
                    {
                        case "DEBUG":
                            model.DebugMode = value == "true";
                            break;
                        case "REPORTNAME":
                            model.ReportName = value;
                            break;
                        case "MSACCESS":
                            model.MsAccess = value;
                            break;
                        case "OUTPUTPDF":
                            model.OutputPdf = value;
                            break;
                        case "FILTER":
                            model.Filter = value;
                            break;
                        case "DATA":
                            model = DecodeDataArgument(value, model);
                            break;
                    }

                    model.Arguments.Add(name, value);
                }
                
            });

            return model;
        }

        private static CommandModel DecodeDataArgument(string data, CommandModel model)
        {
            data = DecodeBase64(data);

            var commands = ParseArguments(data);

            return GetModel(commands);
        }

        private static string DecodeBase64(string encodedString)
        {
            byte[] data = Convert.FromBase64String(encodedString);
            
            return Encoding.UTF8.GetString(data);
        }
        private static string[] ParseArguments(string commandLine)
        {
            char[] parmChars = commandLine.ToCharArray();
            bool inQuote = false;
            for (int index = 0; index < parmChars.Length; index++)
            {
                if (parmChars[index] == '"')
                    inQuote = !inQuote;
                if (!inQuote && parmChars[index] == ' ')
                    parmChars[index] = '\n';
            }
            return (new string(parmChars)).Split('\n');
        }
    }
}
