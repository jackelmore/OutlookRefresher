using System;
using Redemption;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace OutlookRefresher
{
    class AdjustMail
    {
        public static int AdjustTimeStamp(string mailbox)
        {
            string ProjectDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            RedemptionLoader.DllLocation64Bit = ProjectDir + @"\redemption64.dll";
            RedemptionLoader.DllLocation32Bit = ProjectDir + @"\redemption.dll";
            if(!File.Exists(RedemptionLoader.DllLocation32Bit) || !File.Exists(RedemptionLoader.DllLocation64Bit))
            {
                Console.WriteLine("redemption64.dll (64-bit) or redemption.dll (32-bit) is missing from EXE directory\nTerminating with exit code -1");
                return -1;
            }

            RDOSession session = RedemptionLoader.new_RDOSession();
            session.Logon(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            var stores = session.Stores; // store == "mailbox" within an Outlook profile
            foreach (RDOStore rdoStore in stores)
            {
                Debug.WriteLine($"Checking rdoStore == {rdoStore.Name}");
                
                if ((rdoStore.Name.ToLower().Contains(mailbox.ToLower())))
                {
                    Console.WriteLine($"Processing Mailbox {rdoStore.Name}");
                    TimeSpan delta = TimeSpan.Zero; // the amount of time to shift
                    RDOFolder IPMRoot = rdoStore.IPMRootFolder;
//                  var IPMRoot = rdoStore.IPMRootFolder;
                    foreach (RDOFolder folder in IPMRoot.Folders) // find newest e-mail in Inbox
                    {
                        Debug.WriteLine($"  Top Level Folder {folder.Name}");
                        if (folder.Name == "Inbox")
                        {
                            Debug.WriteLine($"    Found {folder.Name} - EntryID {folder.EntryID}");
                            DateTime dtNewest = NewestItemInFolder(folder.EntryID, session);
                            delta = DateTime.Now - dtNewest;
                            Debug.WriteLine($"    Newest item in {folder.Name} is {dtNewest}, delta == {delta}");
                        }
                    }
                    Console.Write($"Fast forward Inbox and Sent Items {delta.Days}d {delta.Hours}h {delta.Minutes}m? [Y/N] :");
                    char c = char.ToUpper(Console.ReadKey().KeyChar);
                    Console.WriteLine();
                    if (c == 'Y')
                    {
                        foreach (RDOFolder folder in IPMRoot.Folders) // adjust dates on all items in Inbox and Sent Items (and their subfolders)
                        {
                            Debug.WriteLine($"  Processing Folder {folder.Name}");
                            if (folder.Name == "Inbox" || folder.Name == "Sent Items")
                            {
                                Debug.WriteLine($"    Found {folder.Name} - EntryID {folder.EntryID}");
                                PerformMailFix(folder.EntryID, session, delta);
                            }
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
                newTimeStamp = oldTimeStamp + delta;
                item.ReceivedTime = newTimeStamp;
                item.Save();
                count++;
                Debug.WriteLine($"      ReceivedTime is {oldTimeStamp}, set to {newTimeStamp}");
            }
            foreach (RDOFolder subFolder in folder.Folders)
            {
                Debug.WriteLine($"      Processing subfolder {subFolder.Name}");
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
