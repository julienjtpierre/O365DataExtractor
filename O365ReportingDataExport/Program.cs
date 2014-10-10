using CommandLine;
using CommandLine.Text;
using Microsoft.Office365.ReportingWebServiceClient;
using Simple.CredentialManager;
using System;

namespace O365ReportingDataExport
{
    internal class Options
    {
        [Option('c', "credential", Required = true, HelpText = "The name of the generic Windows credential you have created in the Windows Credentials Manager")]
        public string CredentialManagerItemName { get; set; }

        [Option('s', "stream", Required = true, HelpText = "The name of the stream")]
        public string StreamName { get; set; }

        [Option('r', "report", Required = true, HelpText = "The name of the report like stated in the Odata feed")]
        public string ReportName { get; set; }

        [Option('f', "from", Required = false, HelpText = "Optional, date from which you want to fetch the data")]
        public DateTime From { get; set; }

        [Option('t', "to", Required = false, HelpText = "Optional, date until which you want to fetch the data")]
        public DateTime To { get; set; }

        [Option('v', "verbose", DefaultValue = true, HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Credential creds = new Credential(null, null, options.CredentialManagerItemName);
                if (creds.Load())
                {
                    ReportingContext context = new ReportingContext();

                    context.UserName = creds.Username;
                    context.Password = creds.Password;
                    context.FromDateTime = options.From;
                    context.ToDateTime = options.To;
                    context.SetLogger(new DefaultLogger());

                    IReportVisitor visitor = new DefaultReportVisitor();

                    ReportingStream stream = new ReportingStream(context, options.ReportName, options.StreamName);
                    stream.RetrieveData(visitor);
                }
            }
        }
    }
}