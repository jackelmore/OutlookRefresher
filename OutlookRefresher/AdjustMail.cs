// OutlookRefresher by Jack Elmore (jackel@microsoft.com)
// Stuff normal people don't try to do to e-mail
// Relies upon Redemption Library http://www.dimastr.com/redemption/home.htm

using System;
using Redemption;
using System.IO;
using System.Reflection;

namespace OutlookRefresher
{
    class AdjustMail
    {
        public static bool backInTime = false;

        public static int AdjustTimeStamp(string mailbox)
        {
            string ProjectDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            RedemptionLoader.DllLocation64Bit = ProjectDir + @"\redemption64.dll";
            RedemptionLoader.DllLocation32Bit = ProjectDir + @"\redemption.dll";
            if (!File.Exists(RedemptionLoader.DllLocation32Bit) || !File.Exists(RedemptionLoader.DllLocation64Bit))
            {
                Console.WriteLine("redemption64.dll (64-bit) or redemption.dll (32-bit) is missing from EXE directory\nTerminating with exit code -1");
                Console.WriteLine($"redemption64.dll should be here: {RedemptionLoader.DllLocation64Bit}");
                Console.WriteLine($"redemption.dll   should be here: {RedemptionLoader.DllLocation32Bit}");
                return -1;
            }

            RDOSession session = RedemptionLoader.new_RDOSession();
            session.Logon(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            var stores = session.Stores; // store == "mailbox" within an Outlook profile
            foreach (RDOStore rdoStore in stores)
            {
                Console.WriteLine($"Checking rdoStore == {rdoStore.Name}");
                
                if ((rdoStore.Name.ToLower().Contains(mailbox.ToLower())))
                {
                    Console.WriteLine($"Processing Mailbox {rdoStore.Name}");
                    TimeSpan delta = TimeSpan.Zero; // the amount of time to shift
                    RDOFolder IPMRoot = rdoStore.IPMRootFolder;
//                  var IPMRoot = rdoStore.IPMRootFolder;
                    foreach (RDOFolder folder in IPMRoot.Folders) // find newest e-mail in Inbox
                    {
                        Console.WriteLine($"  Top Level Folder {folder.Name}");
                        if (folder.Name == "Inbox")
                        {
                            Console.WriteLine($"    Found {folder.Name} - EntryID {folder.EntryID}");
                            DateTime dtNewest = NewestItemInFolder(folder.EntryID, session);
                            delta = DateTime.Now - dtNewest;
                            Console.WriteLine($"    Newest item in {folder.Name} is {dtNewest}, delta == {delta}");
                        }
                    }

                    bool loopUI = true;
                    ConsoleColor OrigConsoleColor = Console.ForegroundColor;
                    while (loopUI)
                    {
                        if(delta > new TimeSpan(100, 0, 0, 0))
                        {

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("\aARE YOU SURE! THIS IS OVER 100 DAYS!");
                            Console.ForegroundColor = OrigConsoleColor;
                            Console.WriteLine();
                        }
                        Console.Write("Adjust Inbox and Sent Items ");
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write(backInTime ? "BACKWARD " : "FORWARD ");
                        Console.Write($"{delta.Days}d {delta.Hours}h {delta.Minutes}m");
                        Console.ForegroundColor = OrigConsoleColor;
                        Console.Write("? [Y]es/[N]o/[C]ustom] :");
                        char c = char.ToUpper(Console.ReadKey().KeyChar);
                        Console.WriteLine();
                        switch (c)
                        {
                            case 'Y':
                                loopUI = false;
                                foreach (RDOFolder folder in IPMRoot.Folders) // adjust dates on all items in Inbox and Sent Items (and their subfolders)
                                {
                                    Console.WriteLine($"  Processing Folder {folder.Name}");
                                    if (folder.Name == "Inbox" || folder.Name == "Sent Items")
                                    {
                                        Console.WriteLine($"    Found {folder.Name} - EntryID {folder.EntryID}");
                                        PerformMailFix(folder.EntryID, session, delta);
                                    }
                                }
                                break;
                            case 'C':
                                Console.WriteLine("6.12:32    6 days 12 hours 32 minutes 00 seconds");
                                Console.WriteLine("6:32       8 hours 32 minutes");
                                Console.WriteLine("-6.12:32   GO BACKWARD 6 days 12 hours 32 minutes 00 seconds");
                                Console.Write("Enter Custom Offset [d].[hh]:[mm]  --> ");
                                delta = TimeSpan.Parse(Console.ReadLine());
                                if(delta < TimeSpan.Zero)
                                {
                                    backInTime = true;
                                    delta = delta - delta - delta; // This is a cheap way to get Abs(TimeSpan) without having to access each field of the struct.
                                }
                                else
                                {
                                    backInTime = false;
                                }
                                break;
                            default:
                                loopUI = false;
                                break;
                        }
                    }
                }
            }
            session.Logoff();
            return count;
        }

        static int count = 0;
        private static void PerformMailFix(string folderId, RDOSession session, TimeSpan delta)
        {
            RDOFolder folder = session.GetFolderFromID(folderId);

            if ( (folder.FolderKind == rdoFolderKind.fkSearch) || (delta == TimeSpan.Zero) )
                return;

            DateTime oldTimeStamp, newTimeStamp;

            foreach (RDOMail item in folder.Items)
            {
                oldTimeStamp = item.ReceivedTime;
                newTimeStamp = backInTime ? oldTimeStamp - delta : oldTimeStamp + delta;
                item.ReceivedTime = newTimeStamp;
                if (!Program.testMode)
                {
                    item.Save();
                }
                count++;
                Console.WriteLine($"      ReceivedTime is {oldTimeStamp}, set to {newTimeStamp}");
            }
            foreach (RDOFolder subFolder in folder.Folders)
            {
                Console.WriteLine($"      Processing subfolder {subFolder.Name}");
                PerformMailFix(subFolder.EntryID, session, delta);
            }
        }
        public static DateTime NewestItemInFolder(string folderID, RDOSession session)
        {
            DateTime dt = DateTime.MinValue;

            RDOFolder folder = session.GetFolderFromID(folderID);

            foreach(RDOMail item in folder.Items)
            {
                if(item.ReceivedTime > dt)
                {
                    dt = item.ReceivedTime;
                }
            }
            return dt;
        }
    }
}
