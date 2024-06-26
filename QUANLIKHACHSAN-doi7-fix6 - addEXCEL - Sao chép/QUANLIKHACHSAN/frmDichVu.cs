﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QUANLIKHACHSAN
{
    public partial class frmDichVu : Form
    {
        public frmDichVu()
        {
            InitializeComponent();
        }
        DBQuanLiKhachSanDataContext dbcontext = new DBQuanLiKhachSanDataContext();
        DataGridViewCellMouseEventArgs vt ;
        bool ktThem;
        string macu = "";
        private void frmDichVu_Load(object sender, EventArgs e)
        {
            KhoaMo(true);
            LayNguon();
        }
        public void XoaTrang()
        {
            txtGhiChu.Text = "";txtGiaDichVu.Text = "";
            txtMaDV.Text = "";txtTenDichVu.Text = "";
        }
        public void KhoaMo(bool key)
        {
            dgvDichVu.Enabled = key;
            btThem.Enabled = key;btSua.Enabled = key;btXoa.Enabled = key;btKetThuc.Enabled = key;
            btGhi.Enabled = !key;btKhongGhi.Enabled = !key;

            txtGhiChu.ReadOnly = key;txtGiaDichVu.ReadOnly = key;
            txtMaDV.ReadOnly = key; txtTenDichVu.ReadOnly = key ;
        }
        public void LayNguon()
        {
            var dl = dbcontext.tbl_DichVus.ToList();
            dgvDichVu.DataSource = dl;
        }

        private void dgvDichVu_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvDichVu.RowCount <=0) return;
            
            if(e.RowIndex >= 0)
            {
                vt = e;
                DataGridViewRow row = dgvDichVu.Rows[e.RowIndex];
                txtMaDV.Text = row.Cells[0].Value.ToString();
                txtTenDichVu.Text = row.Cells[1].Value.ToString();
                txtGiaDichVu.Text = row.Cells[2].Value.ToString();
                txtGhiChu.Text = row.Cells[3].Value.ToString();
            }
        }

        private void btThem_Click(object sender, EventArgs e)
        {
            ktThem = true;
            XoaTrang(); KhoaMo(false);
            txtMaDV.Focus();
        }

        private void btSua_Click(object sender, EventArgs e)
        {
            ktThem = false;
            KhoaMo(false);
            macu = txtMaDV.Text;
            txtTenDichVu.Focus();
            txtMaDV.Enabled = false;
            
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (txtMaDV.Text == "") return;
            if(MessageBox.Show($"ban muon xoa dich vu {txtTenDichVu.Text} ?","THONG BAO",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var objXoa = dbcontext.tbl_DichVus.Where(x => x.MaDV == txtMaDV.Text).SingleOrDefault();
                dbcontext.tbl_DichVus.DeleteOnSubmit(objXoa);
                dbcontext.SubmitChanges();
                XoaTrang();
                LayNguon();
            }
        }

        private void btGhi_Click(object sender, EventArgs e)
        {
            if(txtTenDichVu.Text == "")
            {
                MessageBox.Show("chua nhap ten dich vu", "THONG BAO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenDichVu.Focus();
                return;
            }
            
            if (ktThem == true)
            {
                if (dbcontext.tbl_DichVus.FirstOrDefault(x => x.MaDV == txtMaDV.Text) != null)
                {
                    MessageBox.Show("Dich Vu Da Ton Tai. Hay Kiem Tra Lai Ma Dich Vu", "THONG BAO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtMaDV.Focus();
                    return;
                }
                tbl_DichVu objThem = new tbl_DichVu();
                objThem.MaDV = txtMaDV.Text;
                objThem.TenDV = txtTenDichVu.Text;
                objThem.SoTien = decimal.Parse( txtGiaDichVu.Text);
                objThem.GhiChu = txtGhiChu.Text;
                dbcontext.tbl_DichVus.InsertOnSubmit(objThem);
            }
            else
            {
                var objSua = dbcontext.tbl_DichVus.Where(x => x.MaDV == macu).SingleOrDefault();
                objSua.MaDV = txtMaDV.Text;
                objSua.TenDV = txtTenDichVu.Text;
                objSua.SoTien = decimal.Parse( txtGiaDichVu.Text);
                objSua.GhiChu = txtGhiChu.Text;
            }
            dbcontext.SubmitChanges();
            KhoaMo(true);
            LayNguon();
            txtMaDV.Enabled = true;
        }
        private void btKhongGhi_Click(object sender, EventArgs e)
        {
            try
            {
                XoaTrang();
                KhoaMo(true);
                dgvDichVu_CellMouseClick(sender, vt);
                txtMaDV.Enabled = true;
            }
            catch(Exception ex) { }
            
        }
        private void btKetThuc_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        public void ExportFile(DataTable dataTable, string sheetName, string title)
        {
            //Tạo các đối tượng Excel

            Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbooks oBooks;

            Microsoft.Office.Interop.Excel.Sheets oSheets;

            Microsoft.Office.Interop.Excel.Workbook oBook;

            Microsoft.Office.Interop.Excel.Worksheet oSheet;

            //Tạo mới một Excel WorkBook 

            oExcel.Visible = true;

            oExcel.DisplayAlerts = false;

            oExcel.Application.SheetsInNewWorkbook = 1;

            oBooks = oExcel.Workbooks;

            oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));

            oSheets = oBook.Worksheets;

            oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);

            oSheet.Name = sheetName;

            // Tạo phần Tiêu đề
            Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "D1");

            head.MergeCells = true;

            head.Value2 = title;

            head.Font.Bold = true;

            head.Font.Name = "Times New Roman";

            head.Font.Size = "20";

            head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //thong tin hop dong
            Microsoft.Office.Interop.Excel.Range clten = oSheet.get_Range("A2", "A2");

            clten.Value2 = txtTenDichVu.Text;

            // Tạo tiêu đề cột 

            Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A3", "A3");

            cl1.Value2 = "Mã Dịch Vụ";

            cl1.ColumnWidth = 10.0;

            Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B3", "B3");

            cl2.Value2 = "Tên Dịch Vụ";

            cl2.ColumnWidth = 20.0;

            Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C3", "C3");

            cl3.Value2 = "Số Tiền";
            cl3.ColumnWidth = 15.0;

            Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");

            cl4.Value2 = "Ghi Chu";

            cl4.ColumnWidth = 25.5;


            Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A3", "D3");

            rowHead.Font.Bold = true;

            // Kẻ viền

            rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            // Thiết lập màu nền

            rowHead.Interior.ColorIndex = 5;

            rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // Tạo mảng theo datatable

            object[,] arr = new object[dataTable.Rows.Count, dataTable.Columns.Count];

            //Chuyển dữ liệu từ DataTable vào mảng đối tượng

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                DataRow dataRow = dataTable.Rows[row];

                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    arr[row, col] = dataRow[col];
                }
            }

            //Thiết lập vùng điền dữ liệu

            int rowStart = 4;

            int columnStart = 1;

            int rowEnd = rowStart + dataTable.Rows.Count - 1;

            int columnEnd = dataTable.Columns.Count;

            // Ô bắt đầu điền dữ liệu

            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowStart, columnStart];

            // Ô kết thúc điền dữ liệu

            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnEnd];

            // Lấy về vùng điền dữ liệu

            Microsoft.Office.Interop.Excel.Range range = oSheet.get_Range(c1, c2);

            //Điền dữ liệu vào vùng đã thiết lập

            range.Value2 = arr;

            // Kẻ viền

            range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            // Căn giữa cột mã nhân viên

            //Microsoft.Office.Interop.Excel.Range c3 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnStart];

            //Microsoft.Office.Interop.Excel.Range c4 = oSheet.get_Range(c1, c3);

            //Căn giữa cả bảng 
            oSheet.get_Range(c1, c2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }
        private void btXuatExcel_Click(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();

            DataColumn col1 = new DataColumn("Ma Dich Vu");
            DataColumn col2 = new DataColumn("Ten Dich Vu");
            DataColumn col3 = new DataColumn("Gia Tien");
            DataColumn col4 = new DataColumn("Ghi Chu");
            

            dataTable.Columns.Add(col1);
            dataTable.Columns.Add(col2);
            dataTable.Columns.Add(col3);
            dataTable.Columns.Add(col4);
            

            foreach (DataGridViewRow dgvrow in dgvDichVu.Rows)
            {
                DataRow dtrow = dataTable.NewRow();

                dtrow[0] = dgvrow.Cells[0].Value;
                dtrow[1] = dgvrow.Cells[1].Value;
                dtrow[2] = dgvrow.Cells[2].Value;
                dtrow[3] = dgvrow.Cells[3].Value;

                dataTable.Rows.Add(dtrow);
            }

            ExportFile(dataTable, "dichvu", "Danh Sách Dịch Vụ");
        }
    }
}
