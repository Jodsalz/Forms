using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.OutlookApi.Enums;

using NetOffice;
using Outlook = NetOffice.OutlookApi;

namespace WindowsFormsApplication2
{
    public class Outlok
    {
        Outlook._NameSpace outlookNS;
        Outlook.MAPIFolder inboxFolder;
        Outlook.MAPIFolder ausbuchung;
        Outlook.MAPIFolder einbuchung;
        string body;

        public Outlok()
        {

            NetOffice.OutlookApi.Application outlookApplication = new Outlook.Application();
            outlookApplication.NewMailExEvent += new Outlook.Application_NewMailExEventHandler(outlook_newmail);
            outlookNS = outlookApplication.Session;

            inboxFolder = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["Inbox"];
            einbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Einbuchung"];
            ausbuchung = outlookNS.Folders["devdb.mailhandler@gmail.com"].Folders["[Gmail]"].Folders["DevDB"].Folders["Ausbuchung"];
        }

        private void outlook_newmail(string s)
        {

          
            foreach (COMObject item in inboxFolder.Items)
            {
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                if (mailItem != null && mailItem.Subject == "DevDB Einbuchung")
                {
                    mailItem.Move(einbuchung);
                    //body = mailItem.Body;
                }
                else if (mailItem != null && mailItem.Subject == "DevDB Ausbuchung")
                {
                    mailItem.Move(ausbuchung);
                    //body = mailItem.Body;
                }
            }
            

        }
    }
}
