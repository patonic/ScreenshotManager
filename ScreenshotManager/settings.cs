using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace ScreenshotManager
{
    public partial class settings : Form
    {
        SQLiteConnection connection = null;
        string[] keys = new string[] { "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "[", "]", "A", "S", "D", "F", "G", "H", "J", "K", "L", ";", "'", "Z", "X", "C", "V", "B", "N", "M", ",", ".", "/", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12" };
        Microsoft.Win32.RegistryKey regKey;
        bool admin = true;

        public settings(SQLiteConnection connection)
        {
            InitializeComponent();
            this.connection = connection;
        }

        private void settings_Load(object sender, EventArgs e)
        {
            try
            {
                regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\", true);
                if (regKey.GetValue("ScreenshotManager") != null)
                    autoStartCheckBox.Checked = true;
            }
            catch (Exception)
            {
                admin = false;
                autoStartCheckBox.Enabled = false;
            }

            fullScreenshotShiftCheckBox.Checked = Properties.Settings.Default.fullScreenshotShift;
            fullScreenshotCtrlСheckBox.Checked = Properties.Settings.Default.fullScreenshotCtrl;
            fullScreenshotAltСheckBox.Checked = Properties.Settings.Default.fullScreenshotAlt;
            fullScreenshotKeyComboBox.Items.AddRange(keys);
            fullScreenshotKeyComboBox.SelectedIndex = Array.IndexOf(keys, Properties.Settings.Default.fullScreenshotKey);

            activeScreenshotShiftCheckBox.Checked = Properties.Settings.Default.activeScreenshotShift;
            activeScreenshotCtrlCheckBox.Checked = Properties.Settings.Default.activeScreenshotCtrl;
            activeScreenshotAltCheckBox.Checked = Properties.Settings.Default.activeScreenshotAlt;
            activeScreenshotKeyComboBox.Items.AddRange(keys);
            activeScreenshotKeyComboBox.SelectedIndex = Array.IndexOf(keys, Properties.Settings.Default.activeScreenshotKey);

            textShiftCheckBox.Checked = Properties.Settings.Default.textShift;
            textCtrlCheckBox.Checked = Properties.Settings.Default.textCtrl;
            textAltCheckBox.Checked = Properties.Settings.Default.textAlt;
            textKeyComboBox.Items.AddRange(keys);
            textKeyComboBox.SelectedIndex = Array.IndexOf(keys, Properties.Settings.Default.textKey);

            imageFormatComboBox.DataSource = Enum.GetValues(typeof(ScreenCapture.imageFormats));
            imageFormatComboBox.SelectedItem = (ScreenCapture.imageFormats)Properties.Settings.Default.imageFormat;
            jpgQualityNumericUpDown.Value = Properties.Settings.Default.qualityJpeg;
            if ((ScreenCapture.imageFormats)imageFormatComboBox.SelectedItem != ScreenCapture.imageFormats.jpg)
                jpgQualityNumericUpDown.Enabled = false;

            startInTrayCheckBox.Checked = Properties.Settings.Default.startInTray;

            audioShiftCheckBox.Checked = Properties.Settings.Default.audioShift;
            audioCtrlCheckBox.Checked = Properties.Settings.Default.audioCtrl;
            audioAltCheckBox.Checked = Properties.Settings.Default.audioAlt;
            audioKeyComboBox.Items.AddRange(keys);
            audioKeyComboBox.SelectedIndex = Array.IndexOf(keys, Properties.Settings.Default.audioKey);
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void submitButton_Click(object sender, EventArgs e)
        {
            if (admin) 
            {
                if (autoStartCheckBox.Checked)
                    regKey.SetValue("ScreenshotManager", Application.ExecutablePath);
                else
                    regKey.DeleteValue("ScreenshotManager", false);
            }
            Properties.Settings.Default.startInTray = startInTrayCheckBox.Checked;



            Properties.Settings.Default.fullScreenshotShift = fullScreenshotShiftCheckBox.Checked;
            Properties.Settings.Default.fullScreenshotCtrl = fullScreenshotCtrlСheckBox.Checked;
            Properties.Settings.Default.fullScreenshotAlt = fullScreenshotAltСheckBox.Checked;
            Properties.Settings.Default.fullScreenshotKey = keys[fullScreenshotKeyComboBox.SelectedIndex];

            Properties.Settings.Default.activeScreenshotShift = activeScreenshotShiftCheckBox.Checked;
            Properties.Settings.Default.activeScreenshotCtrl = activeScreenshotCtrlCheckBox.Checked;
            Properties.Settings.Default.activeScreenshotAlt = activeScreenshotAltCheckBox.Checked;
            Properties.Settings.Default.activeScreenshotKey = keys[activeScreenshotKeyComboBox.SelectedIndex];

            Properties.Settings.Default.textShift = textShiftCheckBox.Checked;
            Properties.Settings.Default.textCtrl = textCtrlCheckBox.Checked;
            Properties.Settings.Default.textAlt = textAltCheckBox.Checked;
            Properties.Settings.Default.textKey = keys[textKeyComboBox.SelectedIndex];

            Properties.Settings.Default.imageFormat = (int)imageFormatComboBox.SelectedItem;
            Properties.Settings.Default.qualityJpeg = (byte)jpgQualityNumericUpDown.Value;

            Properties.Settings.Default.audioShift = audioShiftCheckBox.Checked;
            Properties.Settings.Default.audioCtrl = audioCtrlCheckBox.Checked;
            Properties.Settings.Default.audioAlt = audioAltCheckBox.Checked;
            Properties.Settings.Default.audioKey = keys[audioKeyComboBox.SelectedIndex];

            Properties.Settings.Default.Save();
            this.Close();
        }

        private void clearDayButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверенны, что хотите удалить старые записи?\r\nПриложение будет перезапущено.", "Сброс настроек", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                SQLiteCommand sellect = new SQLiteCommand("SELECT id, date FROM files", connection);
                SQLiteDataReader reader = sellect.ExecuteReader();
                while (reader.Read())
                {
                    if (DateTime.Now.AddDays(-1 * Int32.Parse(clearDayNumericUpDown.Value.ToString())) > DateTime.Parse(reader["date"].ToString()))
                    {
                        SQLiteCommand del = new SQLiteCommand("DELETE FROM files WHERE id=@id", connection);
                        del.Parameters.AddWithValue("@id", Int32.Parse(reader["id"].ToString()));
                        del.ExecuteNonQuery();
                    }
                }
                SQLiteCommand vacuum = new SQLiteCommand("VACUUM", connection);
                vacuum.ExecuteNonQuery();
                Application.Restart();
            }
        }

        private void clearStorageButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверенны, что хотите сбросить хранилище?\r\nПриложение будет перезапущено.", "Сброс настроек", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                SQLiteCommand del = new SQLiteCommand("DELETE FROM files; DELETE FROM tags; DELETE FROM tagsForFile; VACUUM", connection);
                del.ExecuteNonQuery();
                Application.Restart();
            }
        }

        private void resetSettingsButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверенны, что хотите сбросить настройки?", "Сброс настроек", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes) {
                Properties.Settings.Default.Reset();
                this.Close();
                Application.Restart();
            }
        }

        private void imageFormatComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((ScreenCapture.imageFormats)imageFormatComboBox.SelectedItem == ScreenCapture.imageFormats.jpg)
                jpgQualityNumericUpDown.Enabled = true;
            else
                jpgQualityNumericUpDown.Enabled = false;
        }
    }
}
