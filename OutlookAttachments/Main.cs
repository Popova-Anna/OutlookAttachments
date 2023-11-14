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
        public Main(IAttachmentSaver attachmentSaver, ILogger logger)
        {
            _attachmentSaver = attachmentSaver;
            _logger = logger;
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var startDate = dtpStartDate.Value.Date;
            var endDate = dtpEndDate.Value.Date.AddDays(1).AddSeconds(-1);

            try
            {
                _attachmentSaver.SaveAttachments(startDate, endDate, System.Configuration.ConfigurationManager.AppSettings["path"]);
                _logger.Information("�������� ������� ���������.");
                MessageBox.Show("�������� ������� ���������. \n������ �� ���������� ����� �������� � ����� log.txt", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger.Error($"������ ���������� ��������: {ex.StackTrace}");
                MessageBox.Show($"������ ���������� ��������: {ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }
}