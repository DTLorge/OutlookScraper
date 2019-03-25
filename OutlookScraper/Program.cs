using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OutlookScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            Class1 GetEmail = new Class1();
            GetEmail.GetAllFolders();
            GetEmail.GetEmailAndInsert2();
        }
    }

    class Class1
    {
        SqlConnection connection;
        DataTable DataTable;
        SqlDataAdapter DataAdapter;
        string folderList = string.Empty;
        private List<Folder> FolderList;

        public Class1()
        {
            FolderList = new List<Folder>();
        }

        public void GetEmailAndInsert()
        {
                Application Myoutlook = new Application();
                NameSpace OutlookNS = Myoutlook.GetNamespace("MAPI");
                Folder inboxfolder = OutlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent;
                EnumerateFolders(inboxfolder);

                foreach (Folder Folder in FolderList)
                {
                    
                    foreach (object obj in Folder.Items)
                    {
                        MailItem Item = obj as MailItem;    
                        if (Item != null && Item.ReceivedTime >= DateTime.Now.AddHours(-1))
                        {
                            InsertMailData("Insert Into MailTable(MailTrackerID,Pulled,Received,Subject,Unread,Atachments,SenderEmail,SenderName,FolderPath)" +
                                "VALUES(NewId(),CURRENT_TIMESTAMP,'" + Item.ReceivedTime + "','" + Item.Subject + "','" + Item.UnRead + "','" +
                                Item.Attachments.Count + "','" + Item.SenderEmailAddress + "','" + Item.SenderName + "','" + Folder.FolderPath + "')");
                        }

                    }

                }

        }

        public void GetEmailAndInsert2()
        {
            Application Myoutlook = new Application();
            NameSpace OutlookNS = Myoutlook.Session;
            //Folder inboxfolder = OutlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent;
            //EnumerateFolders(inboxfolder);
            foreach (Store store in OutlookNS.Stores)
            {
                    EnumerateMAPIFolders(store.GetDefaultFolder(OlDefaultFolders.olFolderInbox));
            }
                foreach (Folder Folder in FolderList)
                {


                    foreach (object obj in Folder.Items)
                    {

                    MailItem Item = null;
                    if (obj is MailItem)
                    {
                        Item = obj as MailItem;
                    }
                    
                        if (Item != null && Item.ReceivedTime >= DateTime.Now.AddHours(-1))
                        {
                            InsertMailData("Insert Into MailTable(MailTrackerID,Pulled,Received,Subject,Unread,Atachments,SenderEmail,SenderName,FolderPath)" +
                                "VALUES(NewId(),CURRENT_TIMESTAMP,'" + Item.ReceivedTime + "','" + Item.Subject + "','" + Item.UnRead + "','" +
                                Item.Attachments.Count + "','" + Item.SenderEmailAddress + "','" + Item.SenderName + "','" + Folder.FolderPath + "')");
                        }

                    }

                }
        }

        private void EnumerateFolders(Folder Folder)
        {
            Folders childFolders =
                Folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    FolderList.Add(childFolder);
                    EnumerateFolders(childFolder);
                }
                

            }
            else FolderList.Add(Folder);
        }

        private void EnumerateMAPIFolders(MAPIFolder Folder)
        {
            Folders childFolders =
                Folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    FolderList.Add(childFolder);
                    EnumerateFolders(childFolder);
                }

            }
            else FolderList.Add(Folder as Folder);

        }

        private void InsertMailData(string Query)
        {
                connection = new SqlConnection(@"Data Source=.;Initial Catalog=MailTracker;Integrated Security=True;");
                DataAdapter = new SqlDataAdapter(Query, connection);
                SqlCommandBuilder command = new SqlCommandBuilder(DataAdapter);
                DataTable = new DataTable();
                DataTable.Locale = System.Globalization.CultureInfo.InvariantCulture;
                DataAdapter.Fill(DataTable);
                DataAdapter.Update(DataTable);
        }



        public void Test()
        {
            
            Application Myoutlook = new Application();
            NameSpace ns = Myoutlook.Session;
            List<Store> stores = GetStores();

            foreach(Store store in stores)
            {
                GetAllFolders();
            }

        }


        public List<Store> GetStores()
        {

            NameSpace ns = null;
            Stores stores = null;
            Application Myoutlook = new Application();
            List<Store> storeList = new List<Store>();

            try
            {
                ns = Myoutlook.Session;
                stores = ns.Stores;

                foreach (Store S in stores)
                {

                    if (S != null)
                    {
                        storeList.Add(S);
                        Marshal.ReleaseComObject(S);
                    }
                }
            }
            finally
            {
                if (stores != null)
                    Marshal.ReleaseComObject(stores);
                if (ns != null)
                    Marshal.ReleaseComObject(ns);
            }
            return storeList;
        }


        private void EnumerateFolders2(Folder Folder)
        {
            Folders childFolders =
                Folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    folderList += "||" + childFolder.Name + "|" +childFolder.Items.Count +  Environment.NewLine;
                    EnumerateFolders(childFolder);
                }


            }
            else FolderList.Add(Folder);
        }

        public void GetAllFolders()
        {
            NameSpace ns = null;
            Stores stores = null;
            MAPIFolder rootFolder = null;
            Folders folders = null;
            MAPIFolder folder = null;
            Application Myoutlook = new Application();


            try
            {
                ns = Myoutlook.Session;
                stores = ns.Stores;
                foreach (Store store in stores)
                {
                    rootFolder = store.GetRootFolder();
                    folders = rootFolder.Folders;
                    folderList += store.DisplayName + Environment.NewLine;
                    for (int i = 1; i < folders.Count; i++)
                    {
                        folder = folders[i];
                        folderList += "|" + folder.Name + "||" + folder.Items.Count+ Environment.NewLine;
                        EnumerateFolders2(folder as Folder);
                        if (folder != null)
                            Marshal.ReleaseComObject(folder);
                    }
                }
                Console.Write(folderList);
            }
            finally
            {
                if (folders != null)
                    Marshal.ReleaseComObject(folders);
                if (folders != null)
                    Marshal.ReleaseComObject(folders);
                if (rootFolder != null)
                    Marshal.ReleaseComObject(rootFolder);
                if (stores != null)
                    Marshal.ReleaseComObject(stores);
                if (ns != null)
                    Marshal.ReleaseComObject(ns);
            }
        }


    }

}


