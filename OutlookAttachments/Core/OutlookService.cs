using Microsoft.Office.Interop.Outlook;
using Serilog;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAttachments.Core
{
    internal class OutlookService : IOutlookService
    {
        private readonly Outlook.Application _outlookApp;
        private readonly NameSpace _outlookNamespace;
        private readonly ILogger _logger;
        public OutlookService(ILogger logger)
        {
            _outlookApp = new Outlook.Application();
            _outlookNamespace = _outlookApp.GetNamespace("MAPI");
            _logger = logger;
        }

        public MailItem[] GetInboxItems(DateTime startDate, DateTime endDate)
        {
            Account? selectedAccount = null;
            try
            {
                foreach (Account _account in _outlookNamespace.Accounts)
                {
                    if (_account.CurrentUser.Address == "eis@zms-chita.ru")
                    {
                        _logger.Information($"Успешное нахождение аккаунта: {_account.CurrentUser.Address}");
                        selectedAccount = _account;
                        break;
                    }
                }
                if (selectedAccount == null)
                {
                    _logger.Error("Ошибка. Не найден аккаунт");
                    throw new ArgumentOutOfRangeException(nameof(selectedAccount), "Ошибка. Не найден аккаунт.");
                    
                }
                Folder? inboxFolder = selectedAccount.DeliveryStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
                if (inboxFolder == null)
                {
                    _logger.Error("Ошибка. Не найдена папка");
                    throw new ArgumentOutOfRangeException(nameof(inboxFolder), "Ошибка. Не найдена папка.");
                }
                Items items = inboxFolder.Items;
                items.Sort("[ReceivedTime]", true); // Сортировка по дате получения письма в порядке убывания
                items = items.Restrict($"[ReceivedTime] >= '{startDate:dd/MM/yyyy HH:mm}' AND [ReceivedTime] <= '{endDate:dd/MM/yyyy HH:mm}'");
                return items.Cast<MailItem>().ToArray();
            }
            catch (System.Exception ex)
            {
                _logger.Error($"Ошибка. Либо не найдена папка. Либо не найден аккаунт. {ex.StackTrace}");
                MessageBox.Show($"Ошибка. Либо не найдена папка. Либо не найден аккаунт. {ex.Message}" );
                throw;
            }

        }

        public void SaveAttachment(Attachment attachment, string filePath)
        {
            if (filePath == null)
            {
                _logger.Error("Ошибка. Пустое место для сохранения данных.");
                throw new ArgumentOutOfRangeException(nameof(filePath), "Ошибка. Пустое место для сохранения данных.");
            }
            try
            {
                attachment.SaveAsFile(filePath);
                _logger.Information($"Сохранение вложения успешно. Путь: {filePath}.");
            }
            catch (System.Exception ex)
            {
                _logger.Error($"Ошибка сохранения файлов. {ex.Message}. StackTrace: {ex.StackTrace}");
                MessageBox.Show($"Ошибка сохранения файлов. {ex.Message}");
                throw;
            }
        }
    }
}
