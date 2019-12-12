using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinformTeamProject
{
    public partial class Form1 : Form
    {
        MySqlConnection conn;
        MySqlDataAdapter dataAdapter, dataAdapter2, dataAdapter3;
        DataSet dataSet;
        int selectedRowIndex;
            bool flg = false;

        private void Form1_Load(object sender, EventArgs e)
        {
            string connStr = "server=localhost;port=3306;database=gguldari;uid=root;pwd=dltpdus0516";
            conn = new MySqlConnection(connStr);
            dataSet = new DataSet();

            dataAdapter = new MySqlDataAdapter("SELECT * FROM customer", conn);
            dataAdapter.Fill(dataSet, "customer");
            dataGridView1.DataSource = dataSet.Tables["customer"];

            dataAdapter2 = new MySqlDataAdapter("SELECT * FROM gguldari.order", conn);
            dataAdapter2.Fill(dataSet, "order");
            dataGridView2.DataSource = dataSet.Tables["order"];

            dataAdapter3 = new MySqlDataAdapter("SELECT * FROM restaurant", conn);
            dataAdapter3.Fill(dataSet, "restaurant");
            dataGridView3.DataSource = dataSet.Tables["restaurant"];

            string[] rank = { "여왕개미", "병정개미", "일개미" };
            cbrank.Items.AddRange(rank);
        }

        public Form1()
        {
            InitializeComponent();
        }

        #region customer 테이블

        private void Btn_Select_Click(object sender, EventArgs e)
        {
            string queryStr;

            #region Select QueryString 만들기
            string[] conditions = new string[10];
            conditions[0] = (textBoxid.Text != "") ? "CustomerId=@id" : null;
            conditions[1] = (textBoxname.Text != "") ? "CustomerName=@name" : null;
            string condition_age;
            if (textBoxage1.Text != "" && textBoxage2.Text != "")
            {
                condition_age = "CustomerAge>=@min and CustomerAge<=@max";
            }
            else if (textBoxage1.Text != "" || textBoxage2.Text != "")
            {
                if (textBoxage1.Text != "")
                    condition_age = "CustomerAge>=@min";
                else
                    condition_age = "CustomerAge<= @max";
            }
            else
            {
                condition_age = null;
            }
            conditions[2] = condition_age;
            conditions[3] = (textBoxjob.Text != "") ? "CustomerJob=@job" : null;
            string condition_point;
            if (textBoxpoint1.Text != "" && textBoxpoint2.Text != "")
            {
                condition_point = "CustomerPoint>=@min1 and CustomerPoint<=@max1";
            }
            else if (textBoxpoint1.Text != "" || textBoxpoint2.Text != "")
            {
                if (textBoxpoint1.Text != "")
                    condition_point = "CustomerPoint>=@min1";
                else
                    condition_point = "CustomerPoint<= @max1";
            }
            else
            {
                condition_point = null;
            }
            conditions[4] = condition_point;
            if(flg==true)
            {
                conditions[5] = (datejoin.Text != "") ? "Customerjoindate=@join" : null;
            }
            else
            {
                conditions[5] = null;
            }
            conditions[6] = (cbrank.Text != "") ? "CustomerRank=@rank" : null;

            if (conditions[0] != null || conditions[1] != null || conditions[2] != null || conditions[3] != null || conditions[4] != null || conditions[5] != null || conditions[6] != null)
            {
                queryStr = $"SELECT * FROM customer WHERE ";
                bool firstCondition = true;
                for (int i = 0; i < conditions.Length; i++)
                {
                    if (conditions[i] != null)
                        if (firstCondition)
                        {
                            queryStr += conditions[i];
                            firstCondition = false;
                        }
                        else
                        {
                            queryStr += " and " + conditions[i];
                        }
                }
            }
            else
            {
                queryStr = "SELECT * FROM customer";
            }
            #endregion

            #region SelectCommand 객체 생성 및 Parameters 설정
            dataAdapter.SelectCommand = new MySqlCommand(queryStr, conn);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@id", textBoxid.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@name", textBoxname.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@min", textBoxage1.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@max", textBoxage2.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@job", textBoxjob.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@min1", textBoxpoint1.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@max1", textBoxpoint2.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@join", datejoin.Text);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@rank", cbrank.Text);
            #endregion

            try
            {
                conn.Open();
                dataSet.Tables["customer"].Clear();
                if (dataAdapter.Fill(dataSet, "customer") > 0)
                    dataGridView1.DataSource = dataSet.Tables["customer"];
                else
                    MessageBox.Show("찾는 데이터가 없습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        internal void UpdateRow(string[] rowDatas)
        {
            string sql = "UPDATE customer SET CustomerName=@name, CustomerAge=@age, CustomerJob=@job, CustomerPoint=@point, Customerjoindate=@join,CustomerRank=@rank WHERE CustomerId=@id";
            dataAdapter.UpdateCommand = new MySqlCommand(sql, conn);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@id", rowDatas[0]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@name", rowDatas[1]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@age", rowDatas[2]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@job", rowDatas[3]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@point", rowDatas[4]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@join", rowDatas[5]);
            dataAdapter.UpdateCommand.Parameters.AddWithValue("@rank", rowDatas[6]);

            try
            {
                conn.Open();
                dataAdapter.UpdateCommand.ExecuteNonQuery();

                dataSet.Tables["customer"].Clear();
                dataAdapter.Fill(dataSet, "customer");
                dataGridView1.DataSource = dataSet.Tables["customer"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void Btn_Insert_Click(object sender, EventArgs e)
        {
            Form2 Dig = new Form2();
            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            textBoxid.Clear();
            textBoxname.Clear();
            textBoxage1.Clear();
            textBoxage2.Clear();
            textBoxjob.Clear();
            textBoxpoint1.Clear();
            textBoxpoint2.Clear();
            datejoin.Text = "";
            checkBox1.Checked = false;
            cbrank.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[selectedRowIndex];

            Form2 Dig = new Form2(
                selectedRowIndex,
                row.Cells[0].Value.ToString(),
                row.Cells[1].Value.ToString(),
                row.Cells[2].Value.ToString(),
                row.Cells[3].Value.ToString(),
                row.Cells[4].Value.ToString(),
                row.Cells[5].Value.ToString(),
                row.Cells[6].Value.ToString()
                );
            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        private void Btn_Save_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("저장할 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // txt 또는 excel 라디오 버튼 상태에 따라 excel 파일로 저장
            if (rbtext.Checked)
            {
                saveFileDialog1.Filter = "텍스트 파일(*.txt)|*.txt";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    SaveTextFile(saveFileDialog1.FileName);
                }
            }
            else
            {
                saveFileDialog1.Filter = "Excel 파일(*.xlsx)|*.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    SaveExcelFile(saveFileDialog1.FileName);

                }
            }

            void SaveTextFile(string filePath)
            {
                // SaveFileDialog에서 지정한 파일경로에 Stream 생성 -> 저장
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                {
                    // Column 이름들 저장
                    foreach (DataColumn col in dataSet.Tables["customer"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataSet의 Row를 저장
                    foreach (DataRow row in dataSet.Tables["customer"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += $"{item.ToString()}\t";
                        }
                        sw.WriteLine(rowString);
                    }

                }
            }
        }

        private void SaveExcelFile(string filePath)
        {
            Excel.Application eApp; 
            Excel.Workbook eworkbook;
            Excel.Worksheet eworksheet;

            eApp = new Excel.Application();
            eworkbook = eApp.Workbooks.Add();
            eworksheet = eworkbook.Sheets[1];

            string[,] dataArr;
            int colCount = dataSet.Tables["customer"].Columns.Count + 1;
            int rowCount = dataSet.Tables["customer"].Rows.Count + 1;
            dataArr = new string[rowCount, colCount];

            for (int i = 0; i < dataSet.Tables["customer"].Columns.Count; i++)
            {
                dataArr[0, i] = dataSet.Tables["customer"].Columns[i].ColumnName;
            }

            for (int i = 0; i < dataSet.Tables["customer"].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables["customer"].Columns.Count; j++)
                {
                    dataArr[i + 1, j] = dataSet.Tables["customer"].Rows[i].ItemArray[j].ToString();
                }
            }

            string endCell = $"G{rowCount}";
            eworksheet.get_Range("A1:" + endCell).Value = dataArr;  

            eworkbook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
            eworkbook.Close(false, Type.Missing, Type.Missing);
            eApp.Quit();
        }  

        internal void InsertRow(string[] rowDatas)
        {
            string queryStr = "INSERT INTO customer (CustomerId,CustomerName,CustomerAge, CustomerJob, CustomerPoint, Customerjoindate, CustomerRank) " +
                "VALUES(@id, @name, @age, @job, @point, @join,@rank)";
            dataAdapter.InsertCommand = new MySqlCommand(queryStr, conn);
            dataAdapter.InsertCommand.Parameters.Add("@id", MySqlDbType.VarChar);
            dataAdapter.InsertCommand.Parameters.Add("@name", MySqlDbType.VarChar);
            dataAdapter.InsertCommand.Parameters.Add("@age", MySqlDbType.Int32);
            dataAdapter.InsertCommand.Parameters.Add("@job", MySqlDbType.VarChar);
            dataAdapter.InsertCommand.Parameters.Add("@point", MySqlDbType.Int32);
            dataAdapter.InsertCommand.Parameters.Add("@join", MySqlDbType.VarChar);
            dataAdapter.InsertCommand.Parameters.Add("@rank", MySqlDbType.VarChar);            
       
            dataAdapter.InsertCommand.Parameters["@id"].Value = rowDatas[0];
            dataAdapter.InsertCommand.Parameters["@name"].Value = rowDatas[1];
            dataAdapter.InsertCommand.Parameters["@age"].Value = rowDatas[2];
            dataAdapter.InsertCommand.Parameters["@job"].Value = rowDatas[3];
            dataAdapter.InsertCommand.Parameters["@point"].Value = rowDatas[4];
            dataAdapter.InsertCommand.Parameters["@join"].Value = rowDatas[5];
            dataAdapter.InsertCommand.Parameters["@rank"].Value = rowDatas[6];

            try
            {
                conn.Open();
                dataAdapter.InsertCommand.ExecuteNonQuery();

                dataSet.Tables["customer"].Clear();
                dataAdapter.Fill(dataSet, "customer");
                dataGridView1.DataSource = dataSet.Tables["customer"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        internal void DeleteRow(string id)
        {
            string sql = "DELETE FROM customer WHERE CustomerId=@id";
            dataAdapter.DeleteCommand = new MySqlCommand(sql, conn);
            dataAdapter.DeleteCommand.Parameters.AddWithValue("@id", id);

            try
            {
                conn.Open();
                dataAdapter.DeleteCommand.ExecuteNonQuery();

                dataSet.Tables["customer"].Clear();
                dataAdapter.Fill(dataSet, "customer");
                dataGridView1.DataSource = dataSet.Tables["customer"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }


        #endregion

        #region order 테이블
        private void order_Select_Click(object sender, EventArgs e)
        {
            string queryStr;

            #region Select QueryString 만들기
            string[] conditions = new string[9];
            conditions[0] = (textBoxodnum.Text != "") ? "OrderNum=@odnum" : null;
            conditions[1] = (textBoxodpd.Text != "") ? "OrderProduct=@odpd" : null;
            conditions[2] = (textBoxoddate.Text != "") ? "OrderDate=@oddate" : null;
            conditions[3] = (textBoxodplace.Text != "") ? "OrderPlace=@odplace" : null;
            conditions[4] = (textBoxodcustomer.Text != "") ? "OrderCustomer=@odcustomer" : null;
            conditions[5] = (textBoxodrestaurant.Text != "") ? "OrderRestaurant=@odrestaurant" : null;


            if (conditions[0] != null || conditions[1] != null || conditions[2] != null || conditions[3] != null || conditions[4] != null || conditions[5] != null)
            {
                queryStr = $"SELECT * FROM gguldari.order WHERE ";
                bool firstCondition = true;
                for (int i = 0; i < conditions.Length; i++)
                {
                    if (conditions[i] != null)
                        if (firstCondition)
                        {
                            queryStr += conditions[i];
                            firstCondition = false;
                        }
                        else
                        {
                            queryStr += " and " + conditions[i];
                        }
                }
            }
            else
            {
                queryStr = "SELECT * FROM gguldari.order";
            }
            #endregion

            #region SelectCommand 객체 생성 및 Parameters 설정
            dataAdapter2.SelectCommand = new MySqlCommand(queryStr, conn);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@odnum", textBoxodnum.Text);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@odpd", textBoxodpd.Text);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@oddate", textBoxoddate.Text);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@odplace", textBoxodplace.Text);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@odcustomer", textBoxodcustomer.Text);
            dataAdapter2.SelectCommand.Parameters.AddWithValue("@odrestaurant", textBoxodrestaurant.Text);
            #endregion

            try
            {
                conn.Open();
                dataSet.Tables["order"].Clear();
                if (dataAdapter2.Fill(dataSet, "order") > 0)
                    dataGridView2.DataSource = dataSet.Tables["order"];
                else
                    MessageBox.Show("찾는 데이터가 없습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void order_Clear_Click(object sender, EventArgs e)
        {
            textBoxodnum.Clear();
            textBoxodpd.Clear();
            textBoxoddate.Clear();
            textBoxodplace.Clear();
            textBoxodcustomer.Clear();
            textBoxodrestaurant.Clear();
        }

        private void Order_Insert_Click(object sender, EventArgs e)
        {
            Form3 Dig = new Form3();
            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        internal void InsertRow1(string[] rowDatas)
        {
            string queryStr = "INSERT INTO gguldari.order (OrderNum,OrderProduct,OrderDate, OrderPlace, OrderCustomer, OrderRestaurantCode) " +
                "VALUES(@odnum,@odpd, @oddate, @odplace, @odcustomer, @odrestaurant)";
            dataAdapter2.InsertCommand = new MySqlCommand(queryStr, conn);
            dataAdapter2.InsertCommand.Parameters.Add("@odnum", MySqlDbType.VarChar);
            dataAdapter2.InsertCommand.Parameters.Add("@odpd", MySqlDbType.VarChar);
            dataAdapter2.InsertCommand.Parameters.Add("@oddate", MySqlDbType.Date);
            dataAdapter2.InsertCommand.Parameters.Add("@odplace", MySqlDbType.VarChar);
            dataAdapter2.InsertCommand.Parameters.Add("@odcustomer", MySqlDbType.VarChar);
            dataAdapter2.InsertCommand.Parameters.Add("@odrestaurant", MySqlDbType.VarChar);

            dataAdapter2.InsertCommand.Parameters["@odnum"].Value = rowDatas[0];
            dataAdapter2.InsertCommand.Parameters["@odpd"].Value = rowDatas[1];
            dataAdapter2.InsertCommand.Parameters["@oddate"].Value = rowDatas[2];
            dataAdapter2.InsertCommand.Parameters["@odplace"].Value = rowDatas[3];
            dataAdapter2.InsertCommand.Parameters["@odcustomer"].Value = rowDatas[4];
            dataAdapter2.InsertCommand.Parameters["@odrestaurant"].Value = rowDatas[5];

            try
            {
                conn.Open();
                dataAdapter2.InsertCommand.ExecuteNonQuery();

                dataSet.Tables["order"].Clear();
                dataAdapter2.Fill(dataSet, "order");
                dataGridView2.DataSource = dataSet.Tables["order"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView2.Rows[selectedRowIndex];

            Form3 Dig = new Form3(
                selectedRowIndex,
                row.Cells[0].Value.ToString(),
                row.Cells[1].Value.ToString(),
                row.Cells[2].Value.ToString(),
                row.Cells[3].Value.ToString(),
                row.Cells[4].Value.ToString(),
                row.Cells[5].Value.ToString()
                );

            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        private void order_Save_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount == 0)
            {
                MessageBox.Show("저장할 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (rbtext1.Checked)
            {
                saveFileDialog2.Filter = "텍스트 파일(*.txt)|*.txt";
                if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    SaveTextFile(saveFileDialog2.FileName);
                }
            }
            else
            {
                saveFileDialog2.Filter = "Excel 파일(*.xlsx)|*.xlsx";
                if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    SaveExcelFile1(saveFileDialog2.FileName);

                }
            }

            void SaveTextFile(string filePath)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog2.FileName))
                {
                    foreach (DataColumn col in dataSet.Tables["order"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataSet의 Row를 저장
                    foreach (DataRow row in dataSet.Tables["order"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += $"{item.ToString()}\t";
                        }
                        sw.WriteLine(rowString);
                    }

                }
            }
        }

        internal void UpdateRow1(string[] rowDatas)
        {
            string sql = "UPDATE gguldari.order SET OrderProduct=@odpd, OrderDate=@oddate, OrderPlace=@odplace,OrderCustomer=@odcustomer, OrderRestaurantCode=@odrestaurant WHERE OrderNum=@odnum";
            dataAdapter2.UpdateCommand = new MySqlCommand(sql, conn);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@odnum", rowDatas[0]);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@odpd", rowDatas[1]);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@oddate", rowDatas[2]);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@odplace", rowDatas[3]);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@odcustomer", rowDatas[4]);
            dataAdapter2.UpdateCommand.Parameters.AddWithValue("@odrestaurant", rowDatas[5]);

            try
            {
                conn.Open();
                dataAdapter2.UpdateCommand.ExecuteNonQuery();

                dataSet.Tables["order"].Clear();
                dataAdapter2.Fill(dataSet, "order");
                dataGridView2.DataSource = dataSet.Tables["order"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        internal void DeleteRow1(string num)
        {
            string sql = "DELETE FROM gguldari.order WHERE OrderNum=@odnum";
            dataAdapter2.DeleteCommand = new MySqlCommand(sql, conn);
            dataAdapter2.DeleteCommand.Parameters.AddWithValue("@odnum", num);

            try
            {
                conn.Open();
                dataAdapter2.DeleteCommand.ExecuteNonQuery();

                dataSet.Tables["order"].Clear();
                dataAdapter2.Fill(dataSet, "order");
                dataGridView2.DataSource = dataSet.Tables["order"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        } 

        private void SaveExcelFile1(string filePath)
        {
            Excel.Application eApp;
            Excel.Workbook eworkbook;
            Excel.Worksheet eworksheet;

            eApp = new Excel.Application();
            eworkbook = eApp.Workbooks.Add();
            eworksheet = eworkbook.Sheets[1];

            string[,] dataArr;
            int colCount = dataSet.Tables["order"].Columns.Count + 1;
            int rowCount = dataSet.Tables["order"].Rows.Count + 1;
            dataArr = new string[rowCount, colCount];

            for (int i = 0; i < dataSet.Tables["order"].Columns.Count; i++)
            {
                dataArr[0, i] = dataSet.Tables["order"].Columns[i].ColumnName;
            }

            for (int i = 0; i < dataSet.Tables["order"].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables["order"].Columns.Count; j++)
                {
                    dataArr[i + 1, j] = dataSet.Tables["order"].Rows[i].ItemArray[j].ToString();
                }
            }

            string endCell = $"F{rowCount}";
            eworksheet.get_Range("A1:" + endCell).Value = dataArr;

            eworkbook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
            eworkbook.Close(false, Type.Missing, Type.Missing);
            eApp.Quit();
        }

        #endregion

        #region restaurant 테이블
        private void res_Select_Click(object sender, EventArgs e)
        {
            string queryStr;

            #region Select QueryString 만들기
            string[] conditions = new string[9];
            conditions[0] = (textBoxodnum.Text != "") ? "RestaurantCode=@rescode" : null;
            conditions[1] = (textBoxodpd.Text != "") ? "RestaurantName=@resname" : null;
            conditions[2] = (textBoxoddate.Text != "") ? "RestaurantAddr=@resaddr" : null;


            if (conditions[0] != null || conditions[1] != null || conditions[2] != null)
            {
                queryStr = $"SELECT * FROM restaurant WHERE ";
                bool firstCondition = true;
                for (int i = 0; i < conditions.Length; i++)
                {
                    if (conditions[i] != null)
                        if (firstCondition)
                        {
                            queryStr += conditions[i];
                            firstCondition = false;
                        }
                        else
                        {
                            queryStr += " and " + conditions[i];
                        }
                }
            }
            else
            {
                queryStr = "SELECT * FROM restaurant";
            }
            #endregion

            #region SelectCommand 객체 생성 및 Parameters 설정
            dataAdapter3.SelectCommand = new MySqlCommand(queryStr, conn);
            dataAdapter3.SelectCommand.Parameters.AddWithValue("@rescode", textBoxrescode.Text);
            dataAdapter3.SelectCommand.Parameters.AddWithValue("@resname", textBoxresname.Text);
            dataAdapter3.SelectCommand.Parameters.AddWithValue("@resaddr", textBoxresaddr.Text);
            #endregion

            try
            {
                conn.Open();
                dataSet.Tables["restaurant"].Clear();
                if (dataAdapter3.Fill(dataSet, "restaurant") > 0)
                    dataGridView3.DataSource = dataSet.Tables["restaurant"];
                else
                    MessageBox.Show("찾는 데이터가 없습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void res_Clear_Click(object sender, EventArgs e)
        {
            textBoxrescode.Clear();
            textBoxresname.Clear();
            textBoxresaddr.Clear();
        }

        private void res_Insert_Click(object sender, EventArgs e)
        {
            Form4 Dig = new Form4();
            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        internal void DeleteRow2(object rescode)
        {
            string sql = "DELETE FROM restaurant WHERE RestaurantCode=@rescode";
            dataAdapter3.DeleteCommand = new MySqlCommand(sql, conn);
            dataAdapter3.DeleteCommand.Parameters.AddWithValue("@rescode", rescode);

            try
            {
                conn.Open();
                dataAdapter3.DeleteCommand.ExecuteNonQuery();

                dataSet.Tables["restaurant"].Clear();
                dataAdapter3.Fill(dataSet, "restaurant");
                dataGridView3.DataSource = dataSet.Tables["restaurant"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView3.Rows[selectedRowIndex];

            Form4 Dig = new Form4(
                selectedRowIndex,
                row.Cells[0].Value.ToString(),
                row.Cells[1].Value.ToString(),
                row.Cells[2].Value.ToString()
                );

            Dig.Owner = this;
            Dig.ShowDialog();
            Dig.Dispose();
        }

        private void res_Save_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount == 0)
            {
                MessageBox.Show("저장할 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (rbtext2.Checked)
            {
                saveFileDialog3.Filter = "텍스트 파일(*.txt)|*.txt";
                if (saveFileDialog3.ShowDialog() == DialogResult.OK)
                {
                    SaveTextFile(saveFileDialog3.FileName);
                }
            }
            else
            {
                saveFileDialog3.Filter = "Excel 파일(*.xlsx)|*.xlsx";
                if (saveFileDialog3.ShowDialog() == DialogResult.OK)
                {
                    SaveExcelFile2(saveFileDialog3.FileName);

                }
            }

            void SaveTextFile(string filePath)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog3.FileName))
                {
                    foreach (DataColumn col in dataSet.Tables["restaurant"].Columns)
                    {
                        sw.Write($"{col.ColumnName}\t");
                    }
                    sw.WriteLine();

                    // DataSet의 Row를 저장
                    foreach (DataRow row in dataSet.Tables["restaurant"].Rows)
                    {
                        string rowString = "";
                        foreach (var item in row.ItemArray)
                        {
                            rowString += $"{item.ToString()}\t";
                        }
                        sw.WriteLine(rowString);
                    }

                }
            }
        }

        private void SaveExcelFile2(string filePath)
        {
            Excel.Application eApp;
            Excel.Workbook eworkbook;
            Excel.Worksheet eworksheet;

            eApp = new Excel.Application();
            eworkbook = eApp.Workbooks.Add();
            eworksheet = eworkbook.Sheets[1];

            string[,] dataArr;
            int colCount = dataSet.Tables["restaurant"].Columns.Count + 1;
            int rowCount = dataSet.Tables["restaurant"].Rows.Count + 1;
            dataArr = new string[rowCount, colCount];

            for (int i = 0; i < dataSet.Tables["restaurant"].Columns.Count; i++)
            {
                dataArr[0, i] = dataSet.Tables["restaurant"].Columns[i].ColumnName;
            }

            for (int i = 0; i < dataSet.Tables["restaurant"].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables["restaurant"].Columns.Count; j++)
                {
                    dataArr[i + 1, j] = dataSet.Tables["restaurant"].Rows[i].ItemArray[j].ToString();
                }
            }

            string endCell = $"C{rowCount}";
            eworksheet.get_Range("A1:" + endCell).Value = dataArr;

            eworkbook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);
            eworkbook.Close(false, Type.Missing, Type.Missing);
            eApp.Quit();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (flg == false) flg = true;
            else flg = false;
        }

        internal void UpdateRow2(string[] rowDatas)
        {
            string sql = "UPDATE restaurant SET RestaurantName=@resname, RestaurantAddr=@resaddr WHERE RestaurantCode=@rescode";
            dataAdapter3.UpdateCommand = new MySqlCommand(sql, conn);
            dataAdapter3.UpdateCommand.Parameters.AddWithValue("@rescode", rowDatas[0]);
            dataAdapter3.UpdateCommand.Parameters.AddWithValue("@resname", rowDatas[1]);
            dataAdapter3.UpdateCommand.Parameters.AddWithValue("@resaddr", rowDatas[2]);

            try
            {
                conn.Open();
                dataAdapter3.UpdateCommand.ExecuteNonQuery();

                dataSet.Tables["restaurant"].Clear();
                dataAdapter3.Fill(dataSet, "restaurant");
                dataGridView1.DataSource = dataSet.Tables["restaurant"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        internal void InsertRow2(string[] rowDatas)
        {
            string queryStr = "INSERT INTO restaurant (RestaurantCode, RestaurantName, RestaurantAddr) " +
                "VALUES(@rescode,@resname,@resaddr)";
            dataAdapter3.InsertCommand = new MySqlCommand(queryStr, conn);
            dataAdapter3.InsertCommand.Parameters.Add("@rescode", MySqlDbType.VarChar);
            dataAdapter3.InsertCommand.Parameters.Add("@resname", MySqlDbType.VarChar);
            dataAdapter3.InsertCommand.Parameters.Add("@resaddr", MySqlDbType.VarChar);

            dataAdapter3.InsertCommand.Parameters["@rescode"].Value = rowDatas[0];
            dataAdapter3.InsertCommand.Parameters["@resname"].Value = rowDatas[1];
            dataAdapter3.InsertCommand.Parameters["@resaddr"].Value = rowDatas[2];

            try
            {
                conn.Open();
                dataAdapter3.InsertCommand.ExecuteNonQuery();

                dataSet.Tables["restaurant"].Clear();
                dataAdapter3.Fill(dataSet, "restaurant");
                dataGridView3.DataSource = dataSet.Tables["restaurant"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }


        #endregion
    }
}
