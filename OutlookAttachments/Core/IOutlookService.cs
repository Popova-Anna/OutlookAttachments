using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachments.Core
{
    public interface IOutlookService
    {
        MailItem[] GetInboxItems(DateTime startDate, DateTime endDate);
        void SaveAttachment(Attachment attachment, string filePath);

        List<string> GetMailAccounts();
    }
}
