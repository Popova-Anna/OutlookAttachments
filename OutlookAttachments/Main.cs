using Newtonsoft.Json;
using OutlookAttachments.Core;
using Serilog;
using System.Windows.Forms;

namespace OutlookAttachments
{
    public partial class Main : Form
    {
        private readonly IAttachmentSaver _attachmentSaver;
        private readonly ILogger _logger;
        private string _path;
        public Main(IAttachmentSaver attachmentSaver, ILogger logger, string path)
        {
            _attachmentSaver = attachmentSaver;
            _logger = logger;
            _path = path;   
            InitializeComponent();         
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var startDate = dtpStartDate.Value.Date;
            var endDate = dtpEndDate.Value.Date.AddDays(1).AddSeconds(-1);

            try
            {
                _attachmentSaver.SaveAttachments(startDate, endDate, _path);
                _logger.Information("Вложения успешно сохранены.");
                MessageBox.Show("Вложения успешно сохранены. \nДанные по сохранённым даным хранятся в файле log.txt", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger.Error($"Ошибка сохранения вложения: {ex.StackTrace}");
                MessageBox.Show($"Ошибка сохранения вложения: {ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }
}