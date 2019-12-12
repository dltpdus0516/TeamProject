using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinformTeamProject
{
    public partial class Form4 : Form
    {
        private string rescode;
        private string resname;
        private string resaddr;
        private int selectedRowIndex;

        public Form4()
        {
            InitializeComponent();
        }

        public Form4(int selectedRowIndex, string v1, string v2, string v3)
        {
            InitializeComponent();
            this.selectedRowIndex = selectedRowIndex;
            rescode = v1;
            resname = v2;
            resaddr = v3;
        }

        Form1 mainForm;
        private void Form4_Load(object sender, EventArgs e)
        {
            txtrescode.Text = rescode;
            txtresname.Text = resname;
            txtresaddr.Text = resaddr;

            if (Owner != null)
            {
                mainForm = Owner as Form1;
            }
        }

        private void btnresInsert_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtrescode.Text,
                txtresname.Text,
                txtresaddr.Text};
            mainForm.InsertRow2(rowDatas);
            this.Close();
        }

        private void btnresUpdate_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtrescode.Text,
                txtresname.Text,
                txtresaddr.Text};
            mainForm.UpdateRow2(rowDatas);
            this.Close();
        }

        private void btnresDelete_Click(object sender, EventArgs e)
        {
            mainForm.DeleteRow2(rescode);
            this.Close();
        }

        private void btnresClear_Click(object sender, EventArgs e)
        {
            txtrescode.Clear();
            txtresname.Clear();
            txtresaddr.Clear();
        }
    }
}
