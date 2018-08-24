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
                Console.WriteLine("Usage: RedemptionApp [-t] <mailbox_name_substring> <mailbox_name_substring> ...");
                Console.WriteLine("-t: Test Mode (no changes saved)");
                return;
            }
            foreach (string arg in args)
            {
                switch(arg.ToLower())
                {
                    case "-t":
                        Console.WriteLine("*** TEST MODE ENABLED (NO CHANGES WILL BE WRITTEN) ***");
                        testMode = true;
                        continue;
                    default:              
                        int itemsProcessed = AdjustMail.AdjustTimeStamp(arg);
                        Console.WriteLine($"Done - {itemsProcessed} items processed. Press any key...");
                        Console.ReadKey();
                        break;
                }
            }
        }
    }
}
