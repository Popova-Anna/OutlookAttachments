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

        /// <summary>
        /// Метод для очистки строки от запрещённых символов.
        /// </summary>
        /// <param name="input">Строка для проверки</param>
        /// <returns>Очищенная строка</returns>
        static private string RemoveForbiddenCharacters(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "Пустая тема";
            char[] forbiddenCharacters = { '/', '\\', '?', '%', '*', ':', '|', '"', '<', '>' };
            string sanitizedInput = new(input.Select(c => forbiddenCharacters.Contains(c) ? '_' : c).ToArray());

            return sanitizedInput;
        }

        /// <summary>
        /// Метод получает письма из Outlook, создает папки для сохранения вложений на основе темы письма и даты получения, и сохраняет вложения в соответствующие папки.
        /// </summary>
        /// <param name="startDate">Дата начала выборки</param>
        /// <param name="endDate">Дата окончания выборки</param>
        /// <param name="saveLocation">Путь сохранения</param>
        public void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation)
        {
           
            
            Outlook.MailItem[] mailItems = _outlookService.GetInboxItems(startDate, endDate);

            foreach (var mailItem in mailItems)
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
                   
                    //Создание директорий
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
