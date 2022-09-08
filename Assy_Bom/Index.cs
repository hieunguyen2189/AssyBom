using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
namespace Assy_Bom
{
    public partial class form_index : Form
    {
        public form_index()
        {
            InitializeComponent();
        }

        private void form_index_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void form_index_Load(object sender, EventArgs e)
        {
            dgv_bind();
            cbb_bind();
            defaultValue();
        }
        private void defaultValue()
        {
            txt_standard_ad.Text = txt_standard_md.Text = "10";
            txt_StdPkg_ad.Text = txt_StdPkg_md.Text = "20";
        }
        private void dgv_bind()
        {
            dgv_model.DataSource = getAllModel();
            dgv_model.RowHeadersVisible = false;
            dgv_model.Columns[0].Width = 160;
            dgv_model.Columns[1].Width = 160;
            dgv_model.Columns[4].Width = 50;
            dgv_model.Columns[5].Width = 50;
            dgv_model.Columns[6].Width = 150;
            dgv_model.Columns[7].Width = 200;
            dgv_model.Columns[8].Width = 60;
            dgv_model.Columns[9].Width = 60;
        }
        private void cbb_bind()
        {

            foreach (DataRow dr in getListModel().Rows)
            {
                cbb_model.Items.Add(dr["MODEL_NAME"].ToString());
            }

        }
        DataTable getAllModel()
        {
            DataTable data = new DataTable();
            string query = "SELECT model_name,customer_model,model_desc,sku_model customer_Sku, model_kp customer_kp, model_type_no sku_category,rev sku_abbreviation,part_desc panel_infomation,standard,std_pkg_qty  FROM sfis1.c_model_desc_t";
            using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
            {
                con.Open();
                OracleCommand Command = new OracleCommand(query, con);
                OracleDataAdapter da = new OracleDataAdapter(Command);
                da.Fill(data);
                con.Close();
            }
            return data;
        }

        DataTable getListModel()
        {
            DataTable data = new DataTable();
            string query = "SELECT distinct(model_name) FROM sfis1.c_model_desc_t order by model_name";
            using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
            {
                con.Open();
                OracleCommand Command = new OracleCommand(query, con);
                Command.ExecuteNonQuery();
                OracleDataAdapter da = new OracleDataAdapter(Command);
                da.Fill(data);
                con.Close();
            }
            return data;
        }


        private void btn_bind_Click(object sender, EventArgs e)
        {
            dgv_bind();
        }

        private void dgv_model_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (dgv_model.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                {
                    dgv_model.CurrentRow.Selected = true;
                    txt_modelDesc_md.Text = txt_DesModel_ad.Text = dgv_model.Rows[e.RowIndex].Cells["MODEL_DESC"].FormattedValue.ToString().ToUpper();
                    txt_CusModel_md.Text = txt_CusModel_ad.Text = dgv_model.Rows[e.RowIndex].Cells["CUSTOMER_MODEL"].FormattedValue.ToString().ToUpper();
                    txt_CusSku_ad.Text = txt_CusSku_md.Text = dgv_model.Rows[e.RowIndex].Cells["CUSTOMER_SKU"].FormattedValue.ToString().ToUpper();
                    txt_CusKp_ad.Text = txt_CusKp_md.Text = dgv_model.Rows[e.RowIndex].Cells["CUSTOMER_KP"].FormattedValue.ToString().ToUpper();
                    txt_CateSku_ad.Text = txt_CateSku_md.Text = dgv_model.Rows[e.RowIndex].Cells["SKU_CATEGORY"].FormattedValue.ToString().ToUpper();
                    txt_AbbSku_ad.Text = txt_AbbSku_md.Text = dgv_model.Rows[e.RowIndex].Cells["SKU_ABBREVIATION"].FormattedValue.ToString().ToUpper();
                    txt_PartDesc_md.Text = txt_DesPart_ad.Text = dgv_model.Rows[e.RowIndex].Cells["PANEL_INFOMATION"].FormattedValue.ToString().ToUpper();
                    txt_standard_ad.Text = txt_standard_md.Text = dgv_model.Rows[e.RowIndex].Cells["STANDARD"].FormattedValue.ToString().ToUpper();
                    txt_StdPkg_ad.Text = txt_StdPkg_md.Text = dgv_model.Rows[e.RowIndex].Cells["STD_PKG_QTY"].FormattedValue.ToString().ToUpper();
                    cbb_model.Text = txt_model_ad.Text = dgv_model.Rows[e.RowIndex].Cells["MODEL_NAME"].FormattedValue.ToString().ToUpper();
                }
            }
            else
            {

            }
        }
        private void cbbChange()
        {
            DataTable data = new DataTable();
            string query = "SELECT model_name,customer_model,model_desc,sku_model customer_Sku, model_kp customer_kp, model_type_no sku_category,rev sku_abbreviation,part_desc panel_infomation,standard,std_pkg_qty  FROM sfis1.c_model_desc_t where model_name = '" + cbb_model.Text + "'";
            using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
            {
                con.Open();
                OracleCommand Command = new OracleCommand(query, con);
                Command.ExecuteNonQuery();
                OracleDataAdapter da = new OracleDataAdapter(Command);
                da.Fill(data);
                con.Close();
            }
            dgv_model.DataSource = data;
            dgv_temp.DataSource = data;
            txt_modelDesc_md.Text = dgv_model.Rows[0].Cells["MODEL_DESC"].FormattedValue.ToString().ToUpper();
            txt_CusModel_md.Text = dgv_model.Rows[0].Cells["CUSTOMER_MODEL"].FormattedValue.ToString().ToUpper();
            txt_CusSku_md.Text = dgv_model.Rows[0].Cells["CUSTOMER_SKU"].FormattedValue.ToString().ToUpper();
            txt_CusKp_md.Text = dgv_model.Rows[0].Cells["CUSTOMER_KP"].FormattedValue.ToString().ToUpper();
            txt_CateSku_md.Text = dgv_model.Rows[0].Cells["SKU_CATEGORY"].FormattedValue.ToString().ToUpper();
            txt_AbbSku_md.Text = dgv_model.Rows[0].Cells["SKU_ABBREVIATION"].FormattedValue.ToString().ToUpper();
            txt_PartDesc_md.Text = dgv_model.Rows[0].Cells["PANEL_INFOMATION"].FormattedValue.ToString().ToUpper();
            txt_standard_md.Text = dgv_model.Rows[0].Cells["STANDARD"].FormattedValue.ToString().ToUpper();
            txt_StdPkg_md.Text = dgv_model.Rows[0].Cells["STD_PKG_QTY"].FormattedValue.ToString().ToUpper();

        }
        private void cbb_model_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbbChange();
        }

        public bool checkModelExist(string modelName)
        {
            string exist_flag = "";
            string modelUpper = modelName.ToUpper();
            string sqlstr1 = "SELECT count(*) FROM sfis1.c_model_desc_t where model_name = '" + modelUpper + "'";
            using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
            {
                con.Open();
                OracleCommand Command = new OracleCommand(sqlstr1, con);
                exist_flag = Command.ExecuteScalar().ToString();
                con.Close();
            }
            if (exist_flag != "0")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool intergerCheck(string text1, string text2)
        {
            int result;
            if (!int.TryParse(text1, out result) || !int.TryParse(text2, out result))
            {
                MessageBox.Show("Only interger Allowed");
                return false;
            }
            else if (Convert.ToInt32(text1) <= 0 || Convert.ToInt32(text2) <= 0)
            {
                MessageBox.Show("Number must be positive");
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool checkNullAdd()
        {
            if (txt_model_ad.Text == "" || txt_CusModel_ad.Text == "" || txt_DesModel_ad.Text == "" || txt_CusSku_ad.Text == "" || txt_CusKp_ad.Text == "" || txt_CateSku_ad.Text == "" || txt_DesPart_ad.Text == "" || txt_standard_ad.Text == "" || txt_StdPkg_ad.Text == "")
            {

                return false;
            }
            else
            {
                return true;
            }
        }
        private bool checkNullMd()
        {
            if (txt_CusModel_md.Text == "" || txt_modelDesc_md.Text == "" || txt_CusSku_md.Text == "" || txt_CusKp_md.Text == "" || txt_CateSku_md.Text == "" || txt_PartDesc_md.Text == "" || txt_standard_md.Text == "" || txt_StdPkg_md.Text == "")
            {

                return false;
            }
            else
            {
                return true;
            }
        }
        private void InsertModel()
        {
            try
            {
                string sqlstr1 = "insert into sfis1.c_model_desc_t (model_name,model_serial,model_type,bom_no,customer,customer_model,model_desc,sku_model,model_kp,model_type_no,rev,part_desc,product_desc,standard,std_pkg_qty) values ('" + txt_model_ad.Text + "','ASSY','1','" + txt_model_ad.Text + "','SONYVN','" + txt_CusModel_ad.Text + "','" + txt_DesModel_ad.Text + "','" + txt_CusSku_ad.Text + "','" + txt_CusKp_ad.Text + "','" + txt_CateSku_ad.Text + "','" + txt_AbbSku_ad.Text + "','" + txt_DesPart_ad.Text + "','CTB','" + txt_standard_ad.Text + "','" + txt_StdPkg_ad.Text + "')";
                using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
                {
                    con.Open();
                    OracleCommand Command = new OracleCommand(sqlstr1, con);
                    Command.ExecuteScalar();
                    con.Close();
                }
                MessageBox.Show("Insert Successfull!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_add_Click(object sender, EventArgs e)
        {

            if (checkNullAdd() == false)
            {
                MessageBox.Show("Please fill all infomation!");
            }
            else
            {
                if (checkModelExist(txt_model_ad.Text) == true)
                {
                    MessageBox.Show("" + txt_model_ad.Text + " already exists! Please try again");
                }
                else
                {
                    if (intergerCheck(txt_standard_ad.Text, txt_StdPkg_ad.Text) == true)
                    {
                        InsertModel();
                    }
                }

            }

        }
        private void clean()
        {
            txt_modelDesc_md.Text = txt_DesModel_ad.Text = txt_CusModel_md.Text = txt_CusModel_ad.Text = txt_CusSku_ad.Text = txt_CusSku_md.Text = txt_CusKp_ad.Text = txt_CusKp_md.Text = txt_CateSku_ad.Text = txt_CateSku_md.Text = txt_AbbSku_ad.Text = txt_AbbSku_md.Text = txt_PartDesc_md.Text = txt_DesPart_ad.Text = txt_model_ad.Text = "";
            cbb_model.Text = "";
            defaultValue();
            dgv_bind();
        }
        private void btn_clean_Click(object sender, EventArgs e)
        {
            clean();
        }
        //}
        private void addArray()
        {
            if (checkModelExist(cbb_model.Text) == true)
            {
                if (checkChanged() == true)
                {
                    if (intergerCheck(txt_standard_md.Text, txt_StdPkg_md.Text) == true)
                    {
                        updateModel();
                    }

                }
                else
                {
                    MessageBox.Show("Nothing changed!!");
                }
            }
            else
            {
                MessageBox.Show("Model name is not exists, Please check again!");

            }
        }
        private void InsertChange(string reason, int cell, string newData, string EmpID, string type)
        {
            try
            {
                string query = "INSERT INTO SFISM4.R_ASSY_BOM_MODIFY_RECORD_T (MODEL_NAME,MODIFY_SECTION,OLD_DATA,NEW_DATA,MODIFIER,MODIFY_TYPE,INSERT_TIME) VALUES ('" + cbb_model.Text + "','" + reason + "','" + dgv_model.Rows[0].Cells[cell].FormattedValue.ToString() + "','" + newData + "','" + EmpID + "','" + type + "',sysdate)";
                using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
                {
                    con.Open();
                    OracleCommand Command = new OracleCommand(query, con);
                    Command.ExecuteScalar();
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private bool checkChanged()
        {
            string flag = "";
            string EmpNo = "V0513581";
            if (dgv_model.Rows[0].Cells["CUSTOMER_MODEL"].FormattedValue.ToString().ToUpper() != txt_CusModel_md.Text)
            {
                InsertChange("CUSTOMER_MODEL", 1, txt_CusModel_md.Text, EmpNo, "MODELMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["MODEL_DESC"].FormattedValue.ToString().ToUpper() != txt_modelDesc_md.Text)
            {
                InsertChange("MODEL_DESC", 2, txt_modelDesc_md.Text, EmpNo, "MODELMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["CUSTOMER_SKU"].FormattedValue.ToString().ToUpper() != txt_CusSku_md.Text)
            {
                InsertChange("CUSTOMER_SKU", 3, txt_CusSku_md.Text, EmpNo, "MODELMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["CUSTOMER_KP"].FormattedValue.ToString().ToUpper() != txt_CusKp_md.Text)
            {
                InsertChange("CUSTOMER_KP", 4, txt_CusKp_md.Text, EmpNo, "MODELMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["SKU_CATEGORY"].FormattedValue.ToString().ToUpper() != txt_CateSku_md.Text)
            {
                InsertChange("SKU_CATEGORY", 5, txt_CateSku_md.Text, EmpNo, "MODELMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["SKU_ABBREVIATION"].FormattedValue.ToString().ToUpper() != txt_AbbSku_md.Text)
            {
                InsertChange("SKU_ABBREVIATION", 6, txt_AbbSku_md.Text, EmpNo, "STANDARDMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["PANEL_INFOMATION"].FormattedValue.ToString().ToUpper() != txt_PartDesc_md.Text)
            {
                InsertChange("PANEL_INFOMATION", 7, txt_PartDesc_md.Text, EmpNo, "STDPKPQTYMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["STANDARD"].FormattedValue.ToString().ToUpper() != txt_standard_md.Text)
            {
                InsertChange("STANDARD", 8, txt_standard_md.Text, EmpNo, "STDPKPQTYMODIFY");
                flag = "1";
            }
            if (dgv_model.Rows[0].Cells["STD_PKG_QTY"].FormattedValue.ToString().ToUpper() != txt_StdPkg_md.Text)
            {
                InsertChange("STD_PKG_QTY", 9, txt_StdPkg_md.Text, EmpNo, "STDPKPQTYMODIFY");
                flag = "1";
            }

            if (flag == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void updateModel()
        {

            try
            {
                string query = "UPDATE sfis1.c_model_desc_t set customer_model='" + txt_CusModel_md.Text + "',model_desc='" + txt_modelDesc_md.Text + "',sku_model='" + txt_CusSku_md.Text + "',model_kp='" + txt_CusKp_md.Text + "',model_type_no='" + txt_CateSku_md.Text + "',rev='" + txt_AbbSku_md.Text + "',part_desc='" + txt_PartDesc_md.Text + "',standard='" + txt_standard_md.Text + "',std_pkg_qty='" + txt_StdPkg_md.Text + "' where model_name = '" + cbb_model.Text + "'";
                using (OracleConnection con = new OracleConnection(ConnectionString.ConnTest))
                {
                    con.Open();
                    OracleCommand Command = new OracleCommand(query, con);
                    Command.ExecuteScalar();
                    con.Close();
                }
                MessageBox.Show("Update Successfull!");
                dgv_bind();
                cbbChange();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btn_update_Click(object sender, EventArgs e)
        {
            if (cbb_model.Text != "")
            {
                addArray();
            }
            else
            {
                MessageBox.Show("Model empty, Please select Model_Name to modify!");
            }
            //if (dgv_temp.Rows[0].Cells[0].FormattedValue.ToString().ToUpper() != "")
            //{
            //    MessageBox.Show(dgv_temp.Rows[0].Cells[0].FormattedValue.ToString().ToUpper());
            //}
        }


        private void dgv_model_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgv_model.ReadOnly = true;
        }

        private void qUITToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            saveExcel();
        }
        private void saveExcel()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx |All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string execPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                var filePath = Path.Combine(execPath, "Template.xlsx");
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks.Open(filePath);
                book.SaveAs(saveFileDialog.FileName); 
                book.Close();
                MessageBox.Show("Save successful!, " + saveFileDialog.FileName);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenFileDialogForm();
        }
        public void OpenFileDialogForm()
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Worksheets|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                dataGridView4.DataSource = ImportExceltoDatatable(openFileDialog1.FileName, "Sheet1");
              
            }
        }
        public static DataTable ImportExceltoDatatable(string filePath, string sheetName)
        {

            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                IXLWorksheet workSheet = workBook.Worksheet(1);
                DataTable dt = new DataTable();
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        if (row.Cell(1).GetString() != "" && row.Cell(2).GetString() != "" && row.Cell(3).GetString() != "" && row.Cell(4).GetString() != "" && row.Cell(5).GetString() != "" && row.Cell(6).GetString() != "")  {
                            foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }
                }

                return dt;
            }
        }
    }
}
