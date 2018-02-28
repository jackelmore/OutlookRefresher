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
            if(args.Length == 0)
            {
                Console.WriteLine("Usage: RedemptionApp <mailbox_name_substring> <mailbox_name_substring> ...");
                return;
            }
            foreach (string arg in args)
            {
                int itemsProcessed = AdjustMail.AdjustTimeStamp(arg);
                Console.WriteLine($"Done - {itemsProcessed} items processed. Press any key...");
                Console.ReadKey();
            }
        }
    }
}
