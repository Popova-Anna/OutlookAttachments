using Newtonsoft.Json;
using OutlookAttachments.Core;
using OutlookAttachments.Model;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAttachments
{
    public partial class Setting : Form
    {
        private ConfigData configData;

        public Setting()
        {
            InitializeComponent();
            // Загрузка данных из файла 
            string json = File.ReadAllText("./config.json");
            configData = JsonConvert.DeserializeObject<ConfigData>(json);
            tBPath.Text = configData.Path;
        }


        private void btnSetPath_Click(object sender, EventArgs e)
        {
            try
            {
                // Загрузка данных из файла 
                string json = File.ReadAllText("./config.json");
                configData = JsonConvert.DeserializeObject<ConfigData>(json);

                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Выберите папку";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        configData.Path = folderDialog.SelectedPath;
                        string updatedJson = JsonConvert.SerializeObject(configData);
                        File.WriteAllText("./config.json", updatedJson);
                    }
                    tBPath.Text = configData.Path;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка:" + ex.Message);
            }
        }

    }
}
