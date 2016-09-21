using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Busylight;

namespace OutlookLightHouse
{

    public partial class ThisAddIn
    {
        private enum Result
        {
            success,
            failure,
            unknown
        }

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        void items_ItemAdd(object Item)
        {
            if (Item == null) { return; }

            Outlook.MailItem mail = (Outlook.MailItem)Item;

            if (mail.MessageClass != "IPM.Note" || !IsCheckinResult(mail)) { return; }

            Dance(GetCheckinResult(mail));
        }

        bool IsCheckinResult(Outlook.MailItem mail)
        {
            return mail.Sender.Address.ToLower().Equals("example@email.com") &&
                   mail.Subject.Contains("[EmailTag]");
        }

        Result GetCheckinResult(Outlook.MailItem mail)
        {
            if (mail.Subject.ToLower().Contains("completed successfully"))
            {
                return Result.success;
            }
            else if (mail.Subject.ToLower().Contains("failed"))
            {
                return Result.failure;
            }

            return Result.unknown;
        }

        void Dance(Result result)
        {
            var sdk = new SDK();
            switch (result)
            {
                case Result.success:
                    sdk.Alert(BusylightColor.Green, BusylightSoundClip.FairyTale, BusylightVolume.Low);
                    System.Threading.Thread.Sleep(3000);
                    sdk.Light(BusylightColor.Off);
                    break;
                case Result.failure:
                    sdk.Alert(BusylightColor.Red, BusylightSoundClip.Funky, BusylightVolume.Low);
                    System.Threading.Thread.Sleep(3000);
                    sdk.Light(BusylightColor.Off);
                    break;
                default:
                    break;
            }
            // TODO: Use skype dll to set light to skype status afterwards.
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
