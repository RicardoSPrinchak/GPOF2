using GemBox.Spreadsheet;
using HarfBuzzSharp;
using System.Diagnostics;
using GPOF2.Properties;
using System.Windows.Forms;
using GemBox.Document;
using GemBox.Document.Tables;
using System.Linq;
using Microsoft.VisualBasic.Devices;

namespace GPOF2
{
    public partial class Form1 : Form
    {
        int test;
        public int suporte1;
        public string escola0;
        public string ensino0;
        public string inep0;
        public string CHANGE;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            CHANGE = Settings.Default["CHANGE"].ToString();
            if (CHANGE == "0")
            {
                escola0 = Settings.Default["ESCOLA0"].ToString();
                inep0 = Settings.Default["INEP0"].ToString();
                ensino0 = Settings.Default["ENSINO0"].ToString();
            }
            else if (CHANGE == "1")
            {
                escola0 = Settings.Default["ESCOLA1"].ToString();
                inep0 = Settings.Default["INEP1"].ToString();
                ensino0 = Settings.Default["ENSINO1"].ToString();
            }
             
            ExcelFile workbook = ExcelFile.Load("GPOF.xlsx");
            var worksheet = workbook.Worksheets[0];
           

            worksheet.Cells["A7"].Value = "Professor(a): " + tbNome.Text;
            worksheet.Cells["A9"].Value = "Ano/Série: " + tbAno.Text;
            worksheet.Cells["B9"].Value = "Turma: " + tbTurma.Text;
            worksheet.Cells["C9"].Value = "Turno: " + tbTurno.Text;
            worksheet.Cells["A13"].Value = "COMPONENTE CURRICULAR: " + tbDisc.Text;

            worksheet.Cells["B18"].Value = P1.Text;
            worksheet.Cells["B18"].Row.AutoFit();
            worksheet.Cells["B19"].Row.AutoFit();
            worksheet.Cells["C18"].Value = P2.Text;
            worksheet.Cells["C18"].Row.AutoFit();
            worksheet.Cells["C19"].Row.AutoFit();
            worksheet.Cells["D18"].Value = P3.Text;
            worksheet.Cells["D18"].Row.AutoFit();
            worksheet.Cells["D19"].Row.AutoFit();
            worksheet.Cells["E18"].Value = P4.Text;
            worksheet.Cells["E18"].Row.AutoFit();
            worksheet.Cells["E19"].Row.AutoFit();
            worksheet.Cells["F18"].Value = P5.Text;
            worksheet.Cells["F18"].Row.AutoFit();
            worksheet.Cells["F19"].Row.AutoFit();
            worksheet.Cells["G18"].Value = P6.Text;
            worksheet.Cells["G18"].Row.AutoFit();
            worksheet.Cells["G19"].Row.AutoFit();
            worksheet.Cells["A5"].Value = escola0;
            worksheet.Cells["A6"].Value = "Código do Inep da Escola: " + inep0;
            worksheet.Cells["A8"].Value = "Nível de Ensino: " + ensino0;

            if (rbPeriodo.Checked == true)
            {
                string periodo = cbPeriodo.SelectedItem.ToString();
                if (periodo.Contains("Bimestre"))
                {
                    worksheet.Cells["A11"].Value = "PLANO BIMESTRAL";
                }
                else if (periodo.Contains("Semestre"))
                {
                    worksheet.Cells["A11"].Value = "PLANO SEMESTRAL";
                }
                else if (periodo.Contains("Trimestre"))
                {
                    worksheet.Cells["A11"].Value = "PLANO TRIMESTRAL";
                }
                
                worksheet.Cells["A18"].Value = cbPeriodo.SelectedItem + " / " + DateTime.Now.ToString("yyyy");

            }
            else if (rbMes.Checked == true)
            {
                worksheet.Cells["A11"].Value = "PLANO MENSAL";
                worksheet.Cells["A18"].Value = cbMes.SelectedItem + " / " + DateTime.Now.ToString("yyyy");
            }
            else if (rbData.Checked == true)
            {
                worksheet.Cells["A11"].Value = "PLANO DE AULA";
                worksheet.Cells["A18"].Value = tbData.Text;
            }

            if (cbPeriodo.Text == "" & cbMes.Text == "" & tbData.Text == "")
            {
                MessageBox.Show("Insira um período.");
            }
            else
            {
                if (rbPdf.Checked == true)
                {
                    if (!Directory.Exists("C:\\GPOF"))
                    {
                        Directory.CreateDirectory("C:\\GPOF");
                    }
                    // worksheet.Cells["A12"].Value = "Documento gerado dia " + DateTime.Now.ToString("dd/MM/yy");
                    string adress;
                    adress = "C:\\GPOF\\" + tbNome.Text + " " + tbAno.Text + " " + tbTurno.Text + " " + DateTime.Now.ToString("ddMMyyHHmmsstt") + ".pdf";
                    workbook.Save(adress);


                    MessageBox.Show("Planejamento gerado.");

                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo(adress)
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }

                if (rbExcel.Checked == true)
                {
                    if (!Directory.Exists("C:\\GPOF"))
                    {
                        Directory.CreateDirectory("C:\\GPOF");
                    }
                    worksheet.Cells["A12"].Value = "Documento gerado dia " + DateTime.Now.ToString("dd/MM/yy");
                    string adress;
                    string adress1;
                    adress = "C:\\GPOF\\" + tbNome.Text + " " + tbAno.Text + " " + tbTurno.Text + " " + DateTime.Now.ToString("ddMMyyHHmmsstt") + ".pdf";
                    adress1 = "C:\\GPOF\\" + tbNome.Text + " " + tbAno.Text + " " + tbTurno.Text + " " + DateTime.Now.ToString("ddMMyyHHmmsstt") + ".xlsx";
                    workbook.Save(adress);
                    MessageBox.Show("Planejamento gerado.");

                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo(adress)
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                    workbook.Save(adress1);
                }
            }
        }
        private void rbPeriodo_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPeriodo.Checked == true)
            {
                cbMes.Enabled = false;
                tbData.Enabled = false;
                cbPeriodo.Enabled = true;
            }
        }
        private void rbMes_CheckedChanged(object sender, EventArgs e)
        {
            if (rbMes.Checked == true)
            {
                cbPeriodo.Enabled = false;
                tbData.Enabled = false;
                cbMes.Enabled = true;
            }
        }
        private void rbData_CheckedChanged(object sender, EventArgs e)
        {
            if (rbData.Checked == true)
            {
                cbPeriodo.Enabled = false;
                cbMes.Enabled = false;
                tbData.Enabled = true;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"c:\GPOF");

        }
        private void button2_Click(object sender, EventArgs e)
        {

            var dlg = new OpenFileDialog()
            {
                InitialDirectory = "C:\\GPOF\\",
                Filter = "Excel Worksheets|*.xlsx",
                RestoreDirectory = true
            };
            //User didn't select a file so return a default value
            if (dlg.ShowDialog() != DialogResult.OK)
            {

            }
            else
            {
                var workbook1 = ExcelFile.Load(dlg.FileName);
                var worksheet1 = workbook1.Worksheets[0];
                // file = "C:\\GPOF\\" + dlg.FileName
                tbNome.Text = worksheet1.Cells["A7"].Value.ToString();
                tbNome.Text = tbNome.Text.Replace("Professor(a): ", "");

                tbAno.Text = worksheet1.Cells["A9"].Value.ToString();
                tbAno.Text = tbAno.Text.Replace("Ano/Série: ", "");

                tbTurma.Text = worksheet1.Cells["B9"].Value.ToString();
                tbTurma.Text = tbTurma.Text.Replace("Turma: ", "");

                tbTurno.Text = worksheet1.Cells["C9"].Value.ToString();
                tbTurno.Text = tbTurno.Text.Replace("Turno: ", "");

                tbDisc.Text = worksheet1.Cells["A13"].Value.ToString();
                tbDisc.Text = tbDisc.Text.Replace("COMPONENTE CURRICULAR: ", "");

                P1.Text = worksheet1.Cells["B18"].Value.ToString();
                P2.Text = worksheet1.Cells["C18"].Value.ToString();
                P3.Text = worksheet1.Cells["D18"].Value.ToString();
                P4.Text = worksheet1.Cells["E18"].Value.ToString();
                P5.Text = worksheet1.Cells["F18"].Value.ToString();
                P6.Text = worksheet1.Cells["G18"].Value.ToString();
            }

        }

        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void sobreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Feito pro Ricardo Souto Princhak, atualmente lotado na E. E. Olga Falcone, email para contato: ricprinchak@hotmail.com");
        }

        private void trocarDadosDaEscolaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 settingForm = new Form2();
            settingForm.ShowDialog();
        }
        private void P1_TextChanged(object sender, EventArgs e)
        {
            
            if (this.P1.Lines.Length > MAX_LINES)
            {
                this.P1.Undo();
                this.P1.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
                
            }
         }
        private const int MAX_LINES = 32;
        private void P1_KeyDown(object sender, KeyEventArgs e)
        {
            if (P1.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
            {
                e.Handled = true;
            }
         }
        private void P2_TextChanged(object sender, EventArgs e)
        {
            
            if (this.P2.Lines.Length > MAX_LINES)
            {
                this.P2.Undo();
                this.P2.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
            }
        }

        private void P2_KeyDown(object sender, KeyEventArgs e)
        {
            if (P2.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
                e.Handled = true;
        }
        private void P3_TextChanged(object sender, EventArgs e)
        {
            if (this.P3.Lines.Length > MAX_LINES)
            {
                this.P3.Undo();
                this.P3.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
            }
        }
        private void P3_KeyDown(object sender, KeyEventArgs e)
        {
            if (P3.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
                e.Handled = true;
        }
        private void P4_TextChanged(object sender, EventArgs e)
        {
            if (this.P4.Lines.Length > MAX_LINES)
            {
                this.P4.Undo();
                this.P4.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
            }
        }
        private void P4_KeyDown(object sender, KeyEventArgs e)
        {
            if(P4.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
                e.Handled = true;
        }
        private void P5_TextChanged(object sender, EventArgs e)
        {
            if (this.P5.Lines.Length > MAX_LINES)
            {
                this.P5.Undo();
                this.P5.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
            }
        }
        private void P5_KeyDown(object sender, KeyEventArgs e)
        {
            if (P5.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
                e.Handled = true;
        }
        private void P6_TextChanged(object sender, EventArgs e)
        {
            if (this.P6.Lines.Length > MAX_LINES)
            {
                this.P6.Undo();
                this.P6.ClearUndo();
                MessageBox.Show("Apenas " + MAX_LINES + " linhas são permitidas.");
            }
        }
        private void P6_KeyDown(object sender, KeyEventArgs e)
        {
            if (P6.Lines.Length >= MAX_LINES && e.KeyValue == '\r')
                e.Handled = true;
        }
        private void label18_Click(object sender, EventArgs e)
        {

        }
    }
}