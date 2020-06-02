using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using NAudio.Wave;
using System.Diagnostics;

namespace ScreenshotManager
{
    public partial class main : Form
    {
        WaveIn waveIn;
        WaveFileWriter writer;
        bool audioWritting = false;

        bool rowBuilding = false;

        #region globalKeyboardHook
        KeyboardHook kh = new KeyboardHook(true);
        private void gkhInstall()
        {
            kh.KeyDown += Kh_KeyDown;
        }

        private void Kh_KeyDown(Keys key, bool Shift, bool Ctrl, bool Alt)
        {
            if ((Keys)Enum.Parse(typeof(Keys), Properties.Settings.Default.fullScreenshotKey.ToUpper()) == key &&
                Ctrl && Alt) {
                createFullScreenshot();
            }

            if ((Keys)Enum.Parse(typeof(Keys), Properties.Settings.Default.activeScreenshotKey.ToUpper()) == key &&
                Ctrl && Alt) {
                activeScreenshot();
            }

            if ((Keys)Enum.Parse(typeof(Keys), Properties.Settings.Default.textKey.ToUpper()) == key &&
                Ctrl && Alt) {
                addText();
            }

            if ((Keys)Enum.Parse(typeof(Keys), Properties.Settings.Default.audioKey.ToUpper()) == key &&
                Ctrl && Alt) {
                audioTrigger();
            }
        }
        #endregion

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        SQLiteConnection connection;

        public main()
        {
            if (!IsSingleInstance())
            {
                this.Close();
                return;
            }

            //SetProcessDPIAware();
            InitializeComponent();

            connection = new SQLiteConnection("Data Source="+Properties.Settings.Default.dbPath+";Version=3;");
            connection.Open();

            if (!Directory.Exists(@"content")) {
                Directory.CreateDirectory(@"content");
            }

            deleteUselessFiles();
        }

        private void main_Load(object sender, EventArgs e)
        {
            loadList();
            gkhInstall();
        }

        private void main_Shown(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.startInTray)
                formToTrey(true);
        }

        #region toolStrip
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void settingsToolStripButton_Click(object sender, EventArgs e)
        {
            settings form = new settings(connection);
            form.ShowDialog();
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count == 0)
                return;

            switch (dataGridView.SelectedRows[0].Cells["type"].Value.ToString())
            {
                case "0": {
                        using (Image img = Image.FromFile(dataGridView.SelectedRows[0].Cells["path"].Value.ToString()))
                        {
                            Clipboard.SetImage(img);
                        }
                    }
                    break;
                case "1":
                    {
                        Clipboard.SetText(string.Join(Environment.NewLine, File.ReadAllLines(dataGridView.SelectedRows[0].Cells["path"].Value.ToString())));
                    }
                    break;
                case "2":
                    {
                        System.Collections.Specialized.StringCollection FileCollection = new System.Collections.Specialized.StringCollection();
                        FileCollection.Add(dataGridView.SelectedRows[0].Cells["path"].Value.ToString());

                        Clipboard.SetFileDropList(FileCollection);
                    }
                    break;
            }
        }

        private void copyPathToolStripButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dataGridView.SelectedRows[0].Cells["path"].Value);
        }
        #endregion

        private void loadList() {
            rowBuilding = true;
            dataGridView.Rows.Clear();
            SQLiteDataReader reader = new SQLiteCommand("SELECT * FROM files ORDER BY id DESC", connection).ExecuteReader();

            while (reader.Read()) {
                if (!File.Exists(@"content\" + reader["name"]))
                {
                    SQLiteCommand command = new SQLiteCommand("DELETE FROM files WHERE id=@id", connection);
                    command.Parameters.AddWithValue("@id", reader["id"]);
                    command.ExecuteNonQuery();
                    return;
                }

                DataGridViewRow row = dataGridView.Rows[dataGridView.Rows.Add()];
                row.Cells["id"].Value = Int32.Parse(reader["id"].ToString());
                row.Cells["date"].Value = DateTime.Parse(reader["date"].ToString()).ToString("dd.MM.yyyy HH:mm:ss");
                row.Cells["path"].Value = Path.GetFullPath(@"content\"+ reader["name"].ToString());
                row.Cells["type"].Value = Int32.Parse(reader["type"].ToString());
                row.Cells["content"].Value = Properties.Settings.Default.textStub;
            }

            rowBuilding = false;
            if (dataGridView.Rows.Count > 0) {
                dataGridView.CurrentCell = null;
                dataGridView.CurrentCell = dataGridView.Rows[0].Cells["content"];
            }

            GC.Collect(1, GCCollectionMode.Forced);
        }

        static bool IsSingleInstance()
        {
            bool flag;
            Mutex mutex = new Mutex(true, "ScreenshotManager_ewlmeymw1216", out flag);
            return flag;
        }

        private void main_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized) {
                formToTrey(true);
            }
        }

        private void notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            formToTrey(false);
        }

        private void formToTrey(bool turnOn) {
            if (turnOn)
            {
                this.Hide();
                this.WindowState = FormWindowState.Minimized;
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    dataGridView.Rows[i].Visible = false;
                }
                dataGridView.Rows.Clear();
                GC.Collect(1, GCCollectionMode.Forced);
            }
            else {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                loadList();
            }
        }

        private void createFullScreenshot() {
            string name = DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss") + " - " + (Int32.Parse((new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar()).ToString()) + 1) + "." + ((ScreenCapture.imageFormats)Properties.Settings.Default.imageFormat).ToString();

            ScreenCapture.CaptureDesktop(@"content\" + name, (ScreenCapture.imageFormats)Properties.Settings.Default.imageFormat, Properties.Settings.Default.qualityJpeg);

            SQLiteCommand insert = new SQLiteCommand("INSERT INTO files (date, name, type) VALUES (@date, @name, @type)", connection);
            insert.Parameters.AddWithValue("@date", DateTime.Now);
            
            insert.Parameters.AddWithValue("@name", name);
            insert.Parameters.AddWithValue("@type", 0);
            
            insert.ExecuteNonQuery();
            if (this.WindowState != FormWindowState.Minimized)
                loadList();
            GC.Collect(1, GCCollectionMode.Forced);
        }

        private void activeScreenshot() {
            string name = DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss") + " - " + (Int32.Parse((new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar()).ToString()) + 1) + "." + ((ScreenCapture.imageFormats)Properties.Settings.Default.imageFormat).ToString();

            ScreenCapture.CaptureActiveWindow(@"content\" + name, (ScreenCapture.imageFormats)Properties.Settings.Default.imageFormat, Properties.Settings.Default.qualityJpeg);

            SQLiteCommand insert = new SQLiteCommand("INSERT INTO files (date, name, type) VALUES (@date, @name, @type)", connection);
            insert.Parameters.AddWithValue("@date", DateTime.Now);

            insert.Parameters.AddWithValue("@name", name);
            insert.Parameters.AddWithValue("@type", 0);

            insert.ExecuteNonQuery();
            if (this.WindowState != FormWindowState.Minimized)
                loadList();
            GC.Collect(1, GCCollectionMode.Forced);
        }

        private void addText() {
            textAdd form = new textAdd();
            form.ShowDialog();
            if (form.lines != null) {
                string name = DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss") + " - " + (Int32.Parse((new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar()).ToString()) + 1) + ".txt";

                System.IO.File.WriteAllLines(@"content\" + name, form.lines);

                SQLiteCommand insert = new SQLiteCommand("INSERT INTO files (date, name, type) VALUES (@date, @name, @type)", connection);
                insert.Parameters.AddWithValue("@date", DateTime.Now);

                insert.Parameters.AddWithValue("@name", name);
                insert.Parameters.AddWithValue("@type", 1);

                insert.ExecuteNonQuery();
                if (this.WindowState != FormWindowState.Minimized)
                    loadList();
            }
            form.Dispose();
        }

        private void audioTrigger() {
            audioWritting = !audioWritting;
            if (audioWritting) {
                notifyIcon.Icon = Properties.Resources.rec;
                try
                {
                    string name = DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss") + " - " + (Int32.Parse((new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar()).ToString()) + 1) + ".wav";

                    waveIn = new WaveIn();
                    waveIn.DeviceNumber = 0;
                    waveIn.DataAvailable += waveIn_DataAvailable;
                    waveIn.RecordingStopped += new EventHandler<NAudio.Wave.StoppedEventArgs>(waveIn_RecordingStopped);
                    waveIn.WaveFormat = new WaveFormat(16000, 1);
                    writer = new WaveFileWriter(@"content\" + name, waveIn.WaveFormat);
                    waveIn.StartRecording();

                    SQLiteCommand insert = new SQLiteCommand("INSERT INTO files (date, name, type) VALUES (@date, @name, @type)", connection);
                    insert.Parameters.AddWithValue("@date", DateTime.Now);

                    insert.Parameters.AddWithValue("@name", name);
                    insert.Parameters.AddWithValue("@type", 2);

                    insert.ExecuteNonQuery();
                    if (this.WindowState != FormWindowState.Minimized)
                        loadList();
                }
                catch (Exception ex)
                { MessageBox.Show(ex.Message); }
            } else {
                notifyIcon.Icon = Properties.Resources.favicon;

                if (waveIn != null)
                {
                    waveIn.StopRecording();
                }
            }
        }

        private void refreshTimer_Tick(object sender, EventArgs e)
        {
            /*if (rowBuilding)
                return;
            rowBuilding = true;*/

            try
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    if (dataGridView.Rows[i].Cells["content"].Displayed)
                    {
                        if (dataGridView.Rows[i].Cells["content"].Value != null && dataGridView.Rows[i].Cells["content"].Value.ToString() == Properties.Settings.Default.textStub)
                        {
                            dataGridView.Rows[i].MinimumHeight = 3;
                            switch (dataGridView.Rows[i].Cells["type"].Value)
                            {
                                case 0:
                                    DataGridViewImageCell imageCell = new DataGridViewImageCell();
                                    imageCell.Value = Image.FromFile(dataGridView.Rows[i].Cells["path"].Value.ToString());
                                    imageCell.Description = dataGridView.Rows[i].Cells["path"].Value.ToString();
                                    imageCell.ImageLayout = DataGridViewImageCellLayout.Zoom;
                                    dataGridView.Rows[i].Cells["content"] = imageCell;
                                    break;
                                case 1:
                                    DataGridViewTextBoxCell textBoxCell = new DataGridViewTextBoxCell();
                                    textBoxCell.Value = string.Join(Environment.NewLine, File.ReadAllLines(dataGridView.Rows[i].Cells["path"].Value.ToString()));
                                    dataGridView.Rows[i].Cells["content"] = textBoxCell;
                                    break;
                                case 2:
                                    DataGridViewButtonCell buttonCell = new DataGridViewButtonCell();
                                    buttonCell.Value = "*Открыть в плеере по умолчанию*";
                                    dataGridView.Rows[i].Cells["content"] = buttonCell;
                                    break;
                                default:
                                    dataGridView.Rows[i].Cells["content"].Value = "Неизвестный тип контента";
                                    break;
                            }
                        }
                    }
                    else if (dataGridView.Rows[i].Cells["content"].Value != null && dataGridView.Rows[i].Cells["content"].Value.ToString() != Properties.Settings.Default.textStub)
                    {
                        dataGridView.Rows[i].MinimumHeight = dataGridView.Rows[i].Height;

                        if (Int32.Parse(dataGridView.Rows[i].Cells["type"].Value.ToString()) == 0)
                            (dataGridView.Rows[i].Cells["content"].Value as Image).Dispose();
                        DataGridViewTextBoxCell cell = new DataGridViewTextBoxCell();
                        cell.Value = Properties.Settings.Default.textStub;
                        dataGridView.Rows[i].Cells["content"] = cell;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            rowBuilding = false;
        }

        void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<WaveInEventArgs>(waveIn_DataAvailable), sender, e);
            }
            else
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
            }
        }

        private void waveIn_RecordingStopped(object sender, EventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler(waveIn_RecordingStopped), sender, e);
            }
            else
            {
                waveIn.Dispose();
                waveIn = null;
                writer.Close();
                writer = null;
            }
        }

        private void dataGridView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["type"].Value.ToString()) == 2 && dataGridView.Columns[e.ColumnIndex].Name == "content") {
                if (audioWritting)
                {
                    MessageBox.Show("Завершите запись аудио перед началом воспроизведения");
                    return;
                }

                Process.Start(dataGridView.Rows[e.RowIndex].Cells["path"].Value.ToString());
            }
        }

        private void dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            SQLiteCommand del = new SQLiteCommand("DELETE FROM files WHERE id=@id", connection);
            del.Parameters.AddWithValue("@id", Int32.Parse(e.Row.Cells["id"].Value.ToString()));
            del.ExecuteNonQuery();
            if (e.Row.Cells["content"].Value != null && e.Row.Cells["content"].Value.ToString() != Properties.Settings.Default.textStub && Int32.Parse(e.Row.Cells["type"].Value.ToString()) == 0)
                (e.Row.Cells["content"].Value as Image).Dispose();
        }

        private void deleteUselessFiles() {
            List<string> files = new List<string>();

            SQLiteCommand comm = new SQLiteCommand("SELECT name FROM files", connection);
            SQLiteDataReader reader = comm.ExecuteReader();
            while (reader.Read())
                files.Add(@"content\" + reader["name"].ToString());
            foreach (string item in Directory.GetFiles(@"content\")) {
                try
                {
                    if (!files.Contains(item))
                        File.Delete(item);
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}
