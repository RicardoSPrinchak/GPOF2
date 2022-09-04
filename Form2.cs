using GPOF2.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GPOF2
{
    public partial class Form2 : Form
    {
        public string change;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            change = Settings.Default["CHANGE"].ToString();
            if (change == "0")
            {
                tbEscola.Text = Settings.Default["ESCOLA0"].ToString();
                tbInep.Text = Settings.Default["INEP0"].ToString();
                tbEnsino.Text = Settings.Default["ENSINO0"].ToString();
            }
            else if (change == "1")
            {
                tbEscola.Text = Settings.Default["ESCOLA1"].ToString();
                tbInep.Text = Settings.Default["INEP1"].ToString();
                tbEnsino.Text = Settings.Default["ENSINO1"].ToString();
            }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Settings.Default["CHANGE"] = "1";
            Settings.Default["ESCOLA1"] = tbEscola.Text;
            Settings.Default["INEP1"] = tbInep.Text;
            Settings.Default["ENSINO1"] = tbEnsino.Text;
            Properties.Settings.Default.Save();
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Settings.Default["CHANGE"] = "0";
            tbEscola.Text = Settings.Default["ESCOLA0"].ToString();
            tbInep.Text = Settings.Default["INEP0"].ToString();
            tbEnsino.Text = Settings.Default["ENSINO0"].ToString();
            Properties.Settings.Default.Save();
        }
    }
}
