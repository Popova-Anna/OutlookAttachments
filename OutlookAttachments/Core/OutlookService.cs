using Microsoft.Office.Interop.Outlook;
using Serilog;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAttachments.Core
{
    public interface IOutlookService
    {
        MailItem[] GetInboxItems(DateTime startDate, DateTime endDate);
        void SaveAttachment(Attachment attachment, string filePath);

        List<Account> GetMailAccounts();
    }
    internal class OutlookService : IOutlookService
    {
        private readonly Outlook.Application _outlookApp;
        private readonly NameSpace _outlookNamespace;
        private readonly ILogger _logger;
        public OutlookService(ILogger logger)
        {
            _outlookApp = new Outlook.Application();
            _outlookNamespace = _outlookApp.GetNamespace("MAPI");
            _outlookNamespace.Logon("eis@zms-chita.ru", "f28e4eJp8", false, true);
            _logger = logger;
        }

        public MailItem[] GetInboxItems(DateTime startDate, DateTime endDate)
        {
            Outlook.Account selectedAccount = null;

            foreach (Outlook.Account _account in _outlookNamespace.Accounts)
            {
                if (_account.DisplayName == "eis@zms-chita.ru")
                {
                    selectedAccount = _account;
                    break;
                }
            }
            Outlook.Folder inboxFolder = selectedAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
           // MAPIFolder inboxFolder = _outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
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
                _logger.Error("Ошибка сохранения файлов." + ex.Message);
                MessageBox.Show("" + ex.Message);
            }
        }
        public List<Account> GetMailAccounts()
        {
            List<Account> accounts = new List<Account>();
            foreach (Account account in _outlookNamespace.Session.Accounts)
            {
                accounts.Add(account);
            }
            return accounts;
        }
    }
}
