using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelWriter.Client
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.DataSource = GetNewDataTable();
        }

        private object GetNewDataTable()
        {
            DataTable data = new DataTable();
            data.Columns.Add();
            data.Columns.Add();
            data.Columns.Add();
            return data;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable data = GetDataFromUI();
            string path = GetPathFromUI();
            if(path == null) { return; }

            ClaimRemedi.ExcelWriter.Writer w = new ClaimRemedi.ExcelWriter.Writer(data, path);
            w.SaveAndCloseFile();
        }

        private string GetPathFromUI()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            if(sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            return null;
        }

        private DataTable GetDataFromUI()
        {
            return (DataTable)this.dataGridView1.DataSource;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable data = GetDataFromUI();
            data.Columns.Add($"Column{data.Columns.Count +1}");
        }
    }
}
