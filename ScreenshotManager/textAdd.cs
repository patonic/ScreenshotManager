using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenshotManager
{
    public partial class textAdd : Form
    {
        public string[] lines;

        public textAdd()
        {
            InitializeComponent();
        }

        private void textAdd_FormClosed(object sender, FormClosedEventArgs e)
        {
            lines = textBox.Lines;
        }

        private void textAdd_Load(object sender, EventArgs e)
        {
            //this.WindowState = FormWindowState.Normal;
            this.Activate();
        }
    }
}
