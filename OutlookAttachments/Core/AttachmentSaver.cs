using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace OutlookAttachments.Core
{
    internal class AttachmentSaver : IAttachmentSaver
    {
        private readonly IOutlookService _outlookService;


        public AttachmentSaver(IOutlookService outlookService)
        {
            _outlookService = outlookService;
        }

        static private string RemoveForbiddenCharacters(string input)
        {
            if (input == "" || input == null)
                return "Пустая тема";
            char[] forbiddenCharacters = { '/', '\\', '?', '%', '*', ':', '|', '"', '<', '>' };
            string sanitizedInput = new(input.Select(c => forbiddenCharacters.Contains(c) ? '_' : c).ToArray());

            return sanitizedInput;
        }

        public void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation)
        {
           
            
            Outlook.MailItem[] mailItems = _outlookService.GetInboxItems(startDate, endDate);

            foreach (Outlook.MailItem mailItem in mailItems)
            {
                //Получаем тему письма
                var subject = RemoveForbiddenCharacters( mailItem.Subject);

                // Получаем дату получения письма
                var receivedDate = mailItem.ReceivedTime; 
                var subjectFolder = Path.Combine(saveLocation, subject);

                // Создаем имя папки по дате получения письма
                var dateFolder = receivedDate.ToString("yyyy-MM-dd");
                var attachmentsFolder = Path.Combine(subjectFolder, dateFolder);


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
                        var attachmentFileName = Path.Combine(attachmentsFolder, attachment.FileName);
                        _outlookService.SaveAttachment(attachment, attachmentFileName);
                    }
                }
            }
        }

    }
}
