using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInPOC
{
    using System.Windows.Forms;
    using Microsoft.Office.Interop.Outlook;

    public partial class PretorRibbon
    {
        private void PretorRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook._Application oApp = new Outlook.Application();
            if (oApp.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = oApp.ActiveExplorer().Selection[1];

                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);

                    // Reads the body of the mail in HTML
                    String htmlBody = mailItem.HTMLBody;

                    // Reads the body of the mail in string
                    String Body = mailItem.Body;

                    int attachRestantes = mailItem.Attachments.Count;

                    for (int j = attachRestantes; j >= 1; j--)
                    {
                        var attachementObject = mailItem.Attachments[j];
                        if (attachementObject.BlockLevel == OlAttachmentBlockLevel.olAttachmentBlockLevelNone)
                        {
                            attachementObject.SaveAsFile(@"C:\temp\attachement" + attachementObject.FileName);
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            // Does some other work.
        }
    }
}
