using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NetOffice.OutlookApi.Enums;

using NetOffice;
using Outlook = NetOffice.OutlookApi;

namespace WindowsFormsApplication2
{
    

    public partial class Form1 : Form
    {

        Outlook._NameSpace outlookNS;
        Outlook.MAPIFolder inboxFolder;
        Outlook.MAPIFolder ausbuchung;
        Outlook.MAPIFolder einbuchung;
        string body = "leer";

        public Form1()
        {
            InitializeComponent();
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
                    body = mailItem.Body;
                   // mailItem.Move(einbuchung);                 
                }
                else if (mailItem != null && mailItem.Subject == "DevDB Ausbuchung")
                {
                    body = mailItem.Body;
                   // mailItem.Move(ausbuchung);
                }
            }
            settext(body);
            
        }

        private void settext(string s)
        {
            label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = s;
            });
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
