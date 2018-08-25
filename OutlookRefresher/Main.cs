// OutlookRefresher by Jack Elmore (jackel@microsoft.com)
// Stuff normal people don't try to do to e-mail
// Relies upon Redemption Library http://www.dimastr.com/redemption/home.htm

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Redemption;

namespace OutlookRefresher
{
    class Program
    {
        // global variables for switches because Honey Badger Don't Care and I am lazy.
        public static bool testMode = false;

        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                Console.WriteLine("Usage: OutlookRefresher [-t] <mailbox_name_substring> <mailbox_name_substring> ...");
                Console.WriteLine("-t: Test Mode (no changes saved)");
                return;
            }
            foreach (string arg in args)
            {
                switch(arg.ToLower())
                {
                    case "-t":
                    case "/t":
                        Console.WriteLine("*** TEST MODE ENABLED (NO CHANGES WILL BE WRITTEN) ***");
                        testMode = true;
                        continue;
                    default:              
                        int itemsProcessed = AdjustMail.AdjustTimeStamp(arg);
                        Console.WriteLine($"Done - {itemsProcessed} items processed.");
                        break;
                }
            }
        }
    }
}
