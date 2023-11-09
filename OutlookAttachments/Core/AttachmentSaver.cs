using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace OutlookAttachments.Core
{
    public interface IAttachmentSaver
    {
        void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation);
    }
    internal class AttachmentSaver : IAttachmentSaver
    {
        private readonly IOutlookService _outlookService;


        public AttachmentSaver(IOutlookService outlookService)
        {
            _outlookService = outlookService;
        }


        private string RemoveForbiddenCharacters(string input)
        {            
            char[] forbiddenCharacters = { '/', '\\', '?', '%', '*', ':', '|', '"', '<', '>', '.' };        
            string sanitizedInput = new string(input.Where(c => !forbiddenCharacters.Contains(c)).ToArray());

            return sanitizedInput;
        }
        public void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation)
        {
           
            
            Outlook.MailItem[] mailItems = _outlookService.GetInboxItems(startDate, endDate);

            foreach (Outlook.MailItem mailItem in mailItems)
            {
                //Получаем тему письма
                string subject = RemoveForbiddenCharacters( mailItem.Subject);

                // Получаем дату получения письма
                DateTime receivedDate = mailItem.ReceivedTime; 
                string subjectFolder = Path.Combine(saveLocation, subject);

                // Создаем имя папки по дате получения письма
                string dateFolder = receivedDate.ToString("yyyy-MM-dd");
                string attachmentsFolder = Path.Combine(subjectFolder, dateFolder);


                if (mailItem.Attachments.Count > 0)
                {
                    //Тема
                    if (!Directory.Exists(subjectFolder))
                    {
                        Directory.CreateDirectory(subjectFolder);
                    }

                    //Дата
                    if (!Directory.Exists(attachmentsFolder))
                    {
                        Directory.CreateDirectory(attachmentsFolder);
                    }

                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        string attachmentFileName = Path.Combine(attachmentsFolder, attachment.FileName);
                        _outlookService.SaveAttachment(attachment, attachmentFileName);
                    }
                }
            }
        }

    }
}
