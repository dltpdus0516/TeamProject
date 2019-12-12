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
    public partial class Form3 : Form
    {
        private string num;
        private string orderpd;
        private string oddate;
        private string odplace;
        private string odcustomer;
        private string odrestaurant;
        private int selectedRowIndex;

        public Form3()
        {
            InitializeComponent();
        }

        public Form3(int selectedRowIndex, string v1, string v2, string v3, string v4, string v5, string v6)
        {
            InitializeComponent();
            this.selectedRowIndex = selectedRowIndex;
            num = v1;
            orderpd = v2;
            oddate = v3;
            odplace = v4;
            odcustomer = v5;
            odrestaurant = v6;
        }

        Form1 mainForm;
        private void Form3_Load(object sender, EventArgs e)
        {
            txtodnum.Text = num;
            txtodproduct.Text = orderpd;
            txtoddate.Text = oddate;
            txtodplace.Text = odplace;
            txtodcustomer.Text = odcustomer;
            txtodrestaurant.Text = odrestaurant;

            if (Owner != null)
            {
                mainForm = Owner as Form1;
            }
        }

        private void btnodInsert_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtodnum.Text,
                txtodproduct.Text,
                txtoddate.Text,
                txtodplace.Text,
                txtodcustomer.Text,
                txtodrestaurant.Text};
            mainForm.InsertRow1(rowDatas);
            this.Close();
        }

        private void btnodUpdate_Click(object sender, EventArgs e)
        {
            string[] rowDatas = {
                txtodnum.Text,
                txtodproduct.Text,
                txtoddate.Text,
                txtodplace.Text,
                txtodcustomer.Text,
                txtodrestaurant.Text};
            mainForm.UpdateRow1(rowDatas);
            this.Close();
        }

        private void btnodDelete_Click(object sender, EventArgs e)
        {
            mainForm.DeleteRow1(num);
            this.Close();
        }

        private void btnodClear_Click(object sender, EventArgs e)
        {
            txtodnum.Clear();
            txtodproduct.Clear();
            txtoddate.Clear();
            txtodplace.Clear();
            txtodcustomer.Clear();
            txtodrestaurant.Clear();
        }
    }
}
