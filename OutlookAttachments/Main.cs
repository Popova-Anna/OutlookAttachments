using Newtonsoft.Json;
using OutlookAttachments.Core;
using OutlookAttachments.Model;
using Serilog;
using System.Windows.Forms;

namespace OutlookAttachments
{
    public partial class Main : Form
    {
        private IAttachmentSaver _attachmentSaver;
        private ILogger _logger;
        private ConfigData configData;
        public Main()
        {
            InitializeComponent();
        }

        private void SetSetting()
        {
            Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
                       .WriteTo.File(configData.Path + "/log.txt")
                       .CreateLogger();

            IOutlookService service = new OutlookService(Log.Logger);
            IAttachmentSaver attachmentSaver = new AttachmentSaver(service);
            _attachmentSaver = attachmentSaver;
            _logger = Log.Logger;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Загрузка данных из файла 
                string json = File.ReadAllText("./config.json");
                configData = JsonConvert.DeserializeObject<ConfigData>(json);
                if (string.IsNullOrEmpty(configData.Path))
                {
                    using (var folderDialog = new FolderBrowserDialog())
                    {
                        folderDialog.Description = "Выберите папку";
                        if (folderDialog.ShowDialog() == DialogResult.OK)
                        {
                            configData.Path = folderDialog.SelectedPath;
                            string updatedJson = JsonConvert.SerializeObject(configData);
                            File.WriteAllText("./config.json", updatedJson);
                        }
                    }
                }
                SetSetting();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка:" + ex.Message);
            }

        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            DateTime startDate = dtpStartDate.Value.Date;
            DateTime endDate = dtpEndDate.Value.Date.AddDays(1).AddSeconds(-1);

            try
            {
                _attachmentSaver.SaveAttachments(startDate, endDate, configData.Path);
                _logger.Information("Вложения успешно сохранены.");
                MessageBox.Show("Вложения успешно сохранены. \nДанные по сохранённым даным хранятся в файле log.txt", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger.Information($"Ошибка сохранения вложения: {ex.Message}");
                MessageBox.Show($"Ошибка сохранения вложения: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void setting_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            setting.ShowDialog();
            string json = File.ReadAllText("./config.json");
            configData = JsonConvert.DeserializeObject<ConfigData>(json);
            SetSetting();
        }
    }
}