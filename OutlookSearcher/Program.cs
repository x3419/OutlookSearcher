using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookSearcher
{
    class Program
    {

        static void Usage()
        {
            string usageString = @"
Usage:
    Arguments:
        Required:
            searchterms     Specify a comma deliminated list of searchterms. e.g.searchterms=foo,bar,asdf

 Examples:
        
        Find all files that have the phrase ""password"" in them.
        
            OutlookSearcher.exe searchterms=password,whatever
            OutlookSearcher.exe searchterms=password
";
            Console.WriteLine(usageString);
        }

        static Dictionary<string, string[]> ParseArgs(string[] args)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string[] commaTerms = new string[] { "searchterms" };
            foreach (string arg in args)
            {
                string[] parts = arg.Split("=".ToCharArray(), 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine("[-] Invalid argument format passed (key/value separated by equals): {0}", arg);
                    continue;
                }
                parts[0] = parts[0].ToLower();
                parts[1] = parts[1].ToLower();
                // Verbosity flag won't have an equal sign
                if (commaTerms.Contains(parts[0]))
                {
                    var tmp = parts[1].Split(',');
                    for (int i = 0; i < tmp.Length; i++)
                    {
                        tmp[i] = tmp[i].Trim();
                    }
                    result[parts[0]] = parts[1].Split(',');
                }
                else
                {
                    result[parts[0]] = new string[] { parts[1] };
                }
            }
            return result;
        }

        static bool ValidateArguments(Dictionary<string, string[]> args)
        {
            /*if (!args.ContainsKey("path"))
            {
                return false;
            }*/
            return true;
        }

        static void Main(string[] args)
        {

            int[] folderTypes = new[] { 3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 18, 19, 20, 21, 22, 23, 25, 28, 29 };

            var parsedArgs = ParseArgs(args);

            if (!ValidateArguments(parsedArgs))
            {
                Usage();
                Environment.Exit(1);
            }

            
            foreach (int i in folderTypes)
            {


                try
                {
                    Application outlookApplication = new Application();

                    NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");

                    MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder((OlDefaultFolders)i);

                    Items mailItems = inboxFolder.Items;

                    StringBuilder stringBuilder = new StringBuilder();

                    foreach (MailItem item in mailItems)
                    {

                        foreach (string val in parsedArgs["searchterms"])
                        {
                            if (item.Subject != null && item.Subject.Contains(val) || item.Body.Contains(val))
                            {
                                stringBuilder.AppendLine("From: " + item.SenderEmailAddress);
                                stringBuilder.AppendLine("To: " + item.To);
                                stringBuilder.AppendLine("Subject: " + item.Subject);
                                stringBuilder.AppendLine(item.Body);

                                Console.WriteLine(stringBuilder.ToString());
                            }

                        }
                        Marshal.ReleaseComObject(item);

                    }
                }
                catch
                {
                    continue;
                }
                
            }
           


            


           
        }
    }
}
