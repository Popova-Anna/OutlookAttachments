using OutlookAttachments.Core;
using Serilog;

namespace OutlookAttachments
{
    public partial class Form1 : Form
    {
        private IAttachmentSaver _attachmentSaver;
        private ILogger _logger;
        public Form1()
        {
            InitializeComponent();
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //Выбор места сохранения вложений и лог файла
            using (FolderBrowserDialog folderBrowserDialog = new())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    txtSaveLocation.Text = folderBrowserDialog.SelectedPath;
                    //кофигурация Логирования
                    Log.Logger = new LoggerConfiguration()
                        .MinimumLevel.Debug()
                        .WriteTo.File(folderBrowserDialog.SelectedPath + "/log.txt")
                        .CreateLogger();

                    IOutlookService service = new OutlookService(Log.Logger);
                    IAttachmentSaver attachmentSaver = new AttachmentSaver(service);
                    _attachmentSaver = attachmentSaver;
                    _logger = Log.Logger;
                    btnSave.Enabled = true;
                }
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            DateTime startDate = dtpStartDate.Value.Date;
            DateTime endDate = dtpEndDate.Value.Date.AddDays(1).AddSeconds(-1);
            string saveLocation = txtSaveLocation.Text;
            try
            {
                _attachmentSaver.SaveAttachments(startDate, endDate, saveLocation);
                _logger.Information("Вложения успешно сохранены.");
                MessageBox.Show("Вложения успешно сохранены. \nДанные по сохранённым даным хранятся в файле log.txt", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                _logger.Information($"Ошибка сохранения вложения: {ex.Message}");
                MessageBox.Show($"Ошибка сохранения вложения: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Log.CloseAndFlush();
        }
    }
}