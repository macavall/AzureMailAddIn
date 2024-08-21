using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace AzureMailAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.MAPIFolder customFolder;
        Outlook.MAPIFolder sentItemsFolder;
        Outlook.Items items;

        private readonly string pattern = @"\d{16}";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            sentItemsFolder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

            customFolder = GetFolderByName("00000Cases");

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(OnItemSend);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void OnItemSend(object item, ref bool cancel)
        {
            if (item is Outlook.MailItem mailItem)
            {
                // Example condition: move if the subject contains "Project X"
                // mailItem.Subject.ToLower().Contains(("blahdeeblah").ToLower())
                if (mailItem.Subject != null && Regex.IsMatch(mailItem.Subject, pattern))
                {
                    // Move the mail item to the custom folder after sending
                    mailItem.SaveSentMessageFolder = customFolder;
                }
            }
        }

        void items_ItemAdd(object Item)
        {
            string filter = "blahdeeblah";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                // mail.Subject.ToUpper().Contains(filter.ToUpper())
                if (mail.MessageClass == "IPM.Note" && Regex.IsMatch(mail.Subject, pattern))
                {
                    mail.Move(customFolder);
                    //mail.Move(outlookNameSpace.GetDefaultFolder(
                    //    Microsoft.Office.Interop.Outlook.
                    //    OlDefaultFolders.));
                }
            }

        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        private Outlook.MAPIFolder GetFolderByName(string folderName)
        {
            Outlook.NameSpace outlookNamespace = this.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            foreach (Outlook.MAPIFolder folder in inboxFolder.Folders)
            {
                if (folder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                {
                    return folder;
                }
            }

            return null; // Folder not found
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
