using System;
using System.Windows.Forms;
using System.Drawing.Imaging;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    class Email
    {
        public static void ExecuteEmailStatus(Excel.Application app)
        {
            try
            {
                CreateImages(app);
                SendEmail(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }

        private static void SendEmail(Excel.Application app)
        {
            Outlook.Application outlook = new Outlook.Application();
            Outlook.MailItem email = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

            email.Subject = "DOT Titling Status: " + DateTime.Now.ToString("M/d/yyyy");
            email.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            Outlook.Recipients recipients = email.Recipients;
            Outlook.Recipient recip1 = recipients.Add("hufflepuff@egov.com");
            recip1.Resolve();
            email.Display();

            string file1 = @ThisAddIn.OutputDir + "\\" + "file1.PNG";
            string file2 = @ThisAddIn.OutputDir + "\\" + "file2.PNG";
            string file3 = @ThisAddIn.OutputDir + "\\" + "file3.PNG";
            string file4 = @ThisAddIn.OutputDir + "\\" + "file4.PNG";
            string file5 = @ThisAddIn.OutputDir + "\\" + "file5.PNG";
            string file6 = @ThisAddIn.OutputDir + "\\" + "file6.PNG";

            var attachment1 = email.Attachments.Add(@file1, Outlook.OlAttachmentType.olEmbeddeditem, null, "file1");
            var attachment2 = email.Attachments.Add(@file2, Outlook.OlAttachmentType.olEmbeddeditem, null, "file2");
            var attachment3 = email.Attachments.Add(@file3, Outlook.OlAttachmentType.olEmbeddeditem, null, "file3");
            var attachment4 = email.Attachments.Add(@file4, Outlook.OlAttachmentType.olEmbeddeditem, null, "file4");
            var attachment5 = email.Attachments.Add(@file5, Outlook.OlAttachmentType.olEmbeddeditem, null, "file5");
            var attachment6 = email.Attachments.Add(@file6, Outlook.OlAttachmentType.olEmbeddeditem, null, "file6");

            string imageCid1 = "file1.png@123";
            string imageCid2 = "file2.png@123";
            string imageCid3 = "file3.png@123";
            string imageCid4 = "file4.png@123";
            string imageCid5 = "file5.png@123";
            string imageCid6 = "file6.png@123";

            attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid1);
            attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid2);
            attachment3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid3);
            attachment4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid4);
            attachment5.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid5);
            attachment6.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid6);

            email.HTMLBody = String.Format("<TABLE cellpadding=5><TR><TD><img src=\"cid:{0}\"></TD></TR><TR><TD><img src=\"cid:{1}\"></TD></TR><TR><TD><img src=\"cid:{2}\"></TD></TR><TR><TD><img src=\"cid:{3}\"></TD></TR><TR><TR><TD><img src=\"cid:{4}\"></TD></TR><TR><TR><TD><img src=\"cid:{5}\"></TD></TR></TABLE>", imageCid1, imageCid2, imageCid3, imageCid4, imageCid5, imageCid6);

            //mailItem.Send();
            recip1 = null;
            //recip2 = null;
            recipients = null;
            email = null;
            outlook = null;
        }

        private static void CreateImages(Excel.Application app)
        {
            try
            {
                Excel.Worksheet wsCover = app.Sheets["Cover"];
                wsCover.Select();

                string file1 = @ThisAddIn.OutputDir + "\\" + "file1.PNG";
                string file2 = @ThisAddIn.OutputDir + "\\" + "file2.PNG";
                string file3 = @ThisAddIn.OutputDir + "\\" + "file3.PNG";
                string file4 = @ThisAddIn.OutputDir + "\\" + "file4.PNG";
                string file5 = @ThisAddIn.OutputDir + "\\" + "file5.PNG";
                string file6 = @ThisAddIn.OutputDir + "\\" + "file6.PNG";

                Excel.Range rng;

                rng = app.get_Range("StatusDevelopmentCurrentSprint", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file1, ImageFormat.Png);
                }

                rng = app.get_Range("StatusDevelopmentCurrentRelease", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file2, ImageFormat.Png);
                }

                rng = app.get_Range("StatusDevelopmentNextRelease", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file3, ImageFormat.Png);
                }

                rng = app.get_Range("StatusRequirements1", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file4, ImageFormat.Png);
                }

                rng = app.get_Range("StatusRequirements2", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file5, ImageFormat.Png);
                }

                rng = app.get_Range("StatusRequirements3", Type.Missing);
                rng.Select();
                rng.Copy();
                if (Clipboard.ContainsImage())
                {
                    Clipboard.GetImage().Save(file6, ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
        }
    }
}
