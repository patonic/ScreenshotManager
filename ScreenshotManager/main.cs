using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using NAudio.Wave;
using System.Diagnostics;
using System.Data.Entity.Core.Metadata.Edm;

namespace ScreenshotManager
{
    public partial class main : Form
    {
        WaveIn waveIn;
        WaveFileWriter writer;
        bool audioWritting = false;
        List<Int32> activeTags = new List<int>();


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

            InitializeComponent();

            connection = new SQLiteConnection("Data Source="+Properties.Settings.Default.dbPath+ ";Version=3; foreign keys=true;");
            connection.Open();

            if (!Directory.Exists(@"content")) {
                Directory.CreateDirectory(@"content");
            }

            deleteUselessFiles();
        }

        private void main_Load(object sender, EventArgs e)
        {
            dateToolStripMenuItem.Checked = Properties.Settings.Default.dateRowDisplaying;
            tagsToolStripMenuItem.Checked = Properties.Settings.Default.tagsRowDisplaying;

            gkhInstall();
            if (Properties.Settings.Default.startInTray)
                return;
            loadTreeView();
            loadList();
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

        private void loadTreeView()
        {
            activeTags.Clear();
            treeView.Nodes.Clear();
            treeView.CheckBoxes = true;
            TreeNode node = new TreeNode("Тип");
            node.Expand();
            TreeNode subnode = new TreeNode("Изображение");
            subnode.Tag = 0;
            subnode.Checked = true;
            node.Nodes.Add(subnode);
            subnode = new TreeNode("Текст");
            subnode.Tag = 1;
            subnode.Checked = true;
            node.Nodes.Add(subnode);
            subnode = new TreeNode("Аудио");
            subnode.Tag = 2;
            subnode.Checked = true;
            node.Nodes.Add(subnode);
            treeView.Nodes.Add(node);
            node = new TreeNode("Теги");
            node.Expand();

            SQLiteDataReader reader = new SQLiteCommand("SELECT * FROM tags", connection).ExecuteReader();
            while (reader.Read())
                node.Nodes.Add(reader["name"].ToString()).Tag = Int32.Parse(reader["id"].ToString());
            node.Nodes.Add("Добавить тег");

            treeView.Nodes.Add(node);
        }

        private void loadList() {
            dataGridView.Rows.Clear();
            string command = "SELECT * FROM files WHERE (false";
            foreach (TreeNode item in treeView.Nodes[0].Nodes)
            {
                if (item.Checked)
                    command = string.Format("{0} OR type='{1}'", command, item.Tag);
            }
            command = command + ")";
            foreach (TreeNode item in treeView.Nodes[1].Nodes)
            {
                if (item.Checked)
                    command = string.Format("{0} AND (SELECT COUNT(*) FROM tagsForFile WHERE files.id=tagsForFile.id_files AND tagsForFile.id_tags='{1}')>0", command, item.Tag);
            }
            command = command + " ORDER BY id DESC";

            SQLiteDataReader reader = new SQLiteCommand(command, connection).ExecuteReader();

            while (reader.Read()) {
                if (!File.Exists(@"content\" + reader["name"]))
                {
                    SQLiteCommand deleteCommand = new SQLiteCommand("DELETE FROM files WHERE id=@id", connection);
                    deleteCommand.Parameters.AddWithValue("@id", reader["id"]);
                    deleteCommand.ExecuteNonQuery();
                    return;
                }

                DataGridViewRow row = dataGridView.Rows[dataGridView.Rows.Add()];
                row.Cells["id"].Value = Int32.Parse(reader["id"].ToString());
                row.Cells["date"].Value = DateTime.Parse(reader["date"].ToString()).ToString("dd.MM.yyyy HH:mm:ss");
                row.Cells["path"].Value = Path.GetFullPath(@"content\"+ reader["name"].ToString());
                row.Cells["type"].Value = Int32.Parse(reader["type"].ToString());
                row.Cells["content"].Value = Properties.Settings.Default.textStub;
            }

            foreach (DataGridViewRow item in dataGridView.Rows)
            {
                reader = new SQLiteCommand(string.Format("SELECT tags.name FROM tagsForFile LEFT JOIN tags ON tags.id=tagsForFile.id_tags WHERE tagsForFile.id_files='{0}'", item.Cells["id"].Value), connection).ExecuteReader();
                string tags = "";
                while (reader.Read())
                    tags = string.Format("{0}{1}; ", tags, reader["name"]);
                dataGridView.Rows[item.Index].Cells["tags"].Value = tags;
            }

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
            if (e.Button == MouseButtons.Left)
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
            addTagsForLastRow();
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
            addTagsForLastRow();
            if (this.WindowState != FormWindowState.Minimized)
                loadList();
            GC.Collect(1, GCCollectionMode.Forced);
        }

        private void addText() {
            textAdd form = new textAdd();
            form.ShowDialog();
            if (form.lines != null) {
                string name = DateTime.Now.ToString("dd-MM-yyyy HH-mm-ss") + " - " + (Int32.Parse((new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar()).ToString()) + 1) + ".txt";

                File.WriteAllLines(@"content\" + name, form.lines);

                SQLiteCommand insert = new SQLiteCommand("INSERT INTO files (date, name, type) VALUES (@date, @name, @type)", connection);
                insert.Parameters.AddWithValue("@date", DateTime.Now);

                insert.Parameters.AddWithValue("@name", name);
                insert.Parameters.AddWithValue("@type", 1);
                
                insert.ExecuteNonQuery();
                addTagsForLastRow();
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
                    addTagsForLastRow();
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

        private void addTagsForLastRow()
        {
            int id = int.Parse(new SQLiteCommand("SELECT last_insert_rowid()", connection).ExecuteScalar().ToString());
            foreach (int id_tags in activeTags)
            {
                SQLiteCommand addTagsForFiles = new SQLiteCommand("INSERT INTO tagsForFile (id_files, id_tags) VALUES (@id_files, @id_tags)", connection);
                addTagsForFiles.Parameters.AddWithValue("@id_files", id);
                addTagsForFiles.Parameters.AddWithValue("@id_tags", id_tags);
                addTagsForFiles.ExecuteNonQuery();
            }
        }

        private void refreshTimer_Tick(object sender, EventArgs e)
        {
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

        private void dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if ((Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["type"].Value.ToString()) == 0 || Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["type"].Value.ToString()) == 1) && dataGridView.Columns[e.ColumnIndex].Name == "content")
            {
                Process.Start(dataGridView.Rows[e.RowIndex].Cells["path"].Value.ToString());
            }
            if (Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["type"].Value.ToString()) == 2 && dataGridView.Columns[e.ColumnIndex].Name == "content")
            {
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

        private void treeView_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node.Level == 0 || (treeView.Nodes.Count > 1 && e.Node ==  treeView.Nodes[1].Nodes[treeView.Nodes[1].Nodes.Count-1])) 
                e.Node.HideCheckBox();
            e.DrawDefault = true;
        }

        private void treeView_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label != null && treeView.Nodes.Count > 1 && e.Node == treeView.Nodes[1].Nodes[treeView.Nodes[1].Nodes.Count - 1] && e.Label != "Добавить тег")
            {
                SQLiteCommand command = new SQLiteCommand("INSERT INTO tags(name) VALUES (@name)", connection);
                command.Parameters.AddWithValue("@name", e.Label);
                command.ExecuteNonQuery();
                loadTreeView();
            }
            else if (e.Label != null && treeView.Nodes.Count > 1 && treeView.Nodes[1].Nodes.Contains(e.Node) && e.Label != "Добавить тег")
            {
                SQLiteCommand command = new SQLiteCommand("UPDATE tags SET name=@name WHERE id=@id", connection);
                command.Parameters.AddWithValue("@name", e.Label);
                command.Parameters.AddWithValue("@id", e.Node.Tag);
                command.ExecuteNonQuery();
                loadTreeView();
            }
        }

        private void treeView_BeforeLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Node.Level == 0)
                e.CancelEdit = true;
        }

        private void treeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            e.Node.BeginEdit();
        }

        private void contextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            contextMenuStrip.Items.Clear();
            SQLiteDataReader reader = new SQLiteCommand("SELECT * FROM tags", connection).ExecuteReader();
            while (reader.Read())
            {
                ToolStripMenuItem item = new ToolStripMenuItem(reader["name"].ToString());
                item.Tag = Int32.Parse(reader["id"].ToString());
                if (activeTags.Contains((Int32)item.Tag))
                {
                    item.BackColor = Color.LightGreen;
                }
                else 
                {
                    item.BackColor = Color.White;
                }
                contextMenuStrip.Items.Add(item);
            }
                
        }

        private void contextMenuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (activeTags.Contains((Int32)e.ClickedItem.Tag))
            {
                activeTags.Remove((Int32)e.ClickedItem.Tag);
                e.ClickedItem.BackColor = Color.White;
            }
            else
            {
                activeTags.Add((Int32)e.ClickedItem.Tag);
                e.ClickedItem.BackColor = Color.LightGreen;
            }
        }

        private void treeView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && treeView.Nodes.Count > 1 && treeView.Nodes[1].Nodes.Contains(treeView.SelectedNode))
            {
                SQLiteCommand command = new SQLiteCommand("DELETE FROM tags WHERE id=@id", connection);
                command.Parameters.AddWithValue("@id", treeView.SelectedNode.Tag);
                command.ExecuteNonQuery();
                loadTreeView();
                loadList();
            }
        }

        private void treeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            loadList();
        }

        private void dateToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.dateRowDisplaying = dateToolStripMenuItem.Checked;
            dataGridView.Columns["date"].Visible = dateToolStripMenuItem.Checked;
            Properties.Settings.Default.Save();
        }

        private void tagsToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.tagsRowDisplaying = tagsToolStripMenuItem.Checked;
            dataGridView.Columns["tags"].Visible = tagsToolStripMenuItem.Checked;
            Properties.Settings.Default.Save();
        }

        
    }
}
