using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WasteDepartments
{
    public partial class Frm_Startup : Form
    {
        public Frm_Startup()
        {
            InitializeComponent();
        }


        private void Frm_Startup_Load_1(object sender, EventArgs e)
        {

            this.Show();
            this.ProgBar();
        }

        private void ProgBar()
        {
            var R = new Random();
            this.progressBar1.Value = this.progressBar1.Minimum;

            while (this.progressBar1.Value < this.progressBar1.Maximum)
            {
                Thread.Sleep(500);
                this.progressBar1.Value = R.Next(this.progressBar1.Value + 1, this.progressBar1.Maximum + 1);
            }
        }


        private void timer1_Tick_1(object sender, EventArgs e)
        {
            Form1 FM = new Form1();
            FM.Show();
            timer1.Enabled = false;
            this.Hide();
        }

        private void pnlImage_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
