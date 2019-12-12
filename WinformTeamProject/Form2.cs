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
    public partial class Form2 : Form
    {
        private string id;
        private string name;
        private string age;
        private string job;
        private string point;
        private string email;
        private string rank;
        private int selectedRowIndex;
        private Form1 mainform;

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(int selectedRowIndex, string v1, string v2, string v3, string v4, string v5, string v6, string v7)
        {
            InitializeComponent();
            this.selectedRowIndex = selectedRowIndex;
            id = v1;
            name = v2;
            age = v3;
            job = v4;
            point = v5;
            email = v6;
            rank = v7;
        }

        private void Form2_Load_1(object sender, EventArgs e)
        {
            txtid.Text = id;
            txtname.Text = name;
            txtage.Text = age;
            txtjob.Text = job;
            txtpoint.Text = point;
            txtjoin.Text = email;
            cbrank.Text = rank;


            if (Owner != null)
            {
                mainform = Owner as Form1;
            }

            string[] rank1 = { "여왕개미", "병정개미", "일개미" };
            cbrank.Items.AddRange(rank1);
        }

        private void btnInsert_Click_1(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtid.Text,
                txtname.Text,
                txtage.Text,
                txtjob.Text,
                txtpoint.Text,
                txtjoin.Text,
                cbrank.Text};
            mainform.InsertRow(rowDatas);
            this.Close();
        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtid.Text,
                txtname.Text,
                txtage.Text,
                txtjob.Text,
                txtpoint.Text,
                txtjoin.Text,
                cbrank.Text};
            mainform.UpdateRow(rowDatas);
            this.Close();
        }

        private void btnDelete_Click_1(object sender, EventArgs e)
        {
            mainform.DeleteRow(id);
            this.Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtid.Clear();
            txtname.Clear();
            txtage.Clear();
            txtjob.Clear();
            txtpoint.Clear();
            txtjoin.Clear();
            cbrank.Text = "";
        }
    }
}
