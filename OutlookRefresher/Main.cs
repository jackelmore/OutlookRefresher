using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Redemption;

namespace OutlookRefresher
{
    class Program
    {
        static void Main(string[] args)
        {
            bool testMode = false;
            if(args.Length == 0)
            {
                Console.WriteLine("Usage: RedemptionApp [-t] <mailbox_name_substring> <mailbox_name_substring> ...");
                Console.WriteLine("-t: Test Mode (no changes saved)");
                return;
            }
            foreach (string arg in args)
            {
                if(arg.ToLower() == "-t")
                {
                    Console.WriteLine("*** TEST MODE ENABLED (NO CHANGES WILL BE WRITTEN) ***");
                    testMode = true;
                    continue;
                }
                int itemsProcessed = AdjustMail.AdjustTimeStamp(arg, testMode);
                Console.WriteLine($"Done - {itemsProcessed} items processed. Press any key...");
                Console.ReadKey();
            }
        }
    }
}
