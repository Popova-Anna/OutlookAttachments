using OutlookAttachments.Core;
using Serilog;

namespace OutlookAttachments
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            try
            {
                var logger = new LoggerConfiguration()
                .WriteTo.File(System.Configuration.ConfigurationManager.AppSettings["path"] + "/log.txt")
                .MinimumLevel.Verbose()
                .CreateLogger();
                var service = new OutlookService(logger);
                var attachmentSaver = new AttachmentSaver(service);

                Application.EnableVisualStyles();
                ApplicationConfiguration.Initialize();
                Application.Run(new Main(attachmentSaver, logger, System.Configuration.ConfigurationManager.AppSettings["path"]));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }

        }
    }
}