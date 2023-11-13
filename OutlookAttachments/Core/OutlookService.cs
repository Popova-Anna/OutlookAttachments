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

            foreach (Account _account in _outlookNamespace.Accounts)
            {
                if (_account.DisplayName == "eis@zms-chita.ru")
                {
                    selectedAccount = _account;
                    break;
                }
            }
            Folder? inboxFolder = selectedAccount?.DeliveryStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;
            Items items = inboxFolder.Items;
            items.Sort("[ReceivedTime]", true); // Сортировка по дате получения письма в порядке убывания
            items = items.Restrict($"[ReceivedTime] >= '{startDate:dd/MM/yyyy HH:mm}' AND [ReceivedTime] <= '{endDate:dd/MM/yyyy HH:mm}'");
            return items.Cast<MailItem>().ToArray();
        }


        public void SaveAttachment(Attachment attachment, string filePath)
        {
            try
            {
                attachment.SaveAsFile(filePath);
                _logger.Information("Сохранение вложения успешно. Путь:" + filePath);
            }
            catch (System.Exception ex)
            {
                _logger.Error("Ошибка сохранения файлов." + ex.StackTrace);
                MessageBox.Show("" + ex.Message);
                throw;
            }
        }
        public List<string> GetMailAccounts()
        {
            List<string> accounts = new();
            foreach (Account account in _outlookNamespace.Session.Accounts)
            {
                accounts.Add(account.DisplayName);
            }
            return accounts;
        }
    }
}
