// OutlookRefresher by Jack Elmore (jackel@microsoft.com)
// Stuff normal people don't try to do to e-mail
// Relies upon Redemption Library http://www.dimastr.com/redemption/home.htm

using System;
using System.Diagnostics;

namespace OutlookRefresher
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: OutlookRefresher [-t] [-v] <mailbox_name_substring> <mailbox_name_substring> ...");
                Console.WriteLine("-t: Test Mode (no changes saved)");
                Console.WriteLine("-v: Verbose (Debug.WriteLine output to console)");
                return 1;
            }
            else
            {
                foreach (string arg in args)
                {
                    switch (arg.ToLower())
                    {
                        case "-t":
                        case "/t":
                            Console.WriteLine("******************************************************");
                            Console.WriteLine("*** TEST MODE ENABLED (NO CHANGES WILL BE WRITTEN) ***");
                            Console.WriteLine("******************************************************");
                            Options.TestMode = true;
                            continue;
                        case "-v":
                        case "/v":
                            Console.WriteLine("*** VERBOSE LOGGING ENABLED ***");
                            Options.VerboseLogging = true;
                            Debug.Listeners.Add(new ConsoleTraceListener(useErrorStream: false));
                            break;
                        default:
                            Console.WriteLine("{0} total items adjusted. Exiting...", AdjustMail.AdjustTimeStamp(arg));
                            break;
                    }
                }
            }
            return 0;
        }
    }
}
