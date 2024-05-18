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

namespace QUANLIKHACHSAN
{
    public partial class frmVatTu : Form
    {
        DBQuanLiKhachSanDataContext dbcontext = new DBQuanLiKhachSanDataContext();
        string sDK;
        DataGridViewCellMouseEventArgs vt;
        bool ktThem;
        string macu;
        public frmVatTu()
        {
            InitializeComponent();
        }
        public void XoaTrang()
        {
            txtGhiChu.Text = ""; txtTimKiem.Text = "";
            txtSoLuong.Text = ""; txtTenVatTu.Text = ""; txtGhiChu.Text = "Null";
        }
        public void KhoaMo(bool key)
        {
            dgvVatTu.Enabled = key;
            btThem.Enabled = key; btSua.Enabled = key; btXoa.Enabled = key; btKetThuc.Enabled = key;
            btGhi.Enabled = !key; btKhongGhi.Enabled = !key;

            txtGhiChu.ReadOnly = key; txtSoLuong.ReadOnly = key; txtTenPhong.ReadOnly = key;
            txtMaVatTu.ReadOnly = key; txtTenVatTu.ReadOnly = key; txtTimKiem.ReadOnly = !key;
        }
        public void LayNguonVatTu()
        {
            var dl = (from vtu in dbcontext.tbl_VatTus
                      where ((ckbHienTatCa.Checked == true)
                      && (sDK == "" ? true : vtu.MaPhong.Contains(sDK)))
                      select new
                      {
                          vtu.MaPhong,
                          vtu.MaVatTu,
                          vtu.TenVatTu,
                          vtu.SoLuong,
                          vtu.GhiChu
                      }
                   ).ToList();
            dgvVatTu.DataSource = dl;
        }

        private void frmVatTu_Load(object sender, EventArgs e)
        {
            KhoaMo(true);
            LayNguonVatTu();
        }

        private void ckbHienTatCa_CheckedChanged(object sender, EventArgs e)
        {
            sDK = txtTimKiem.Text;
            LayNguonVatTu();
        }

        private void dgvVatTu_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvVatTu.RowCount <= 0) return;

            if (e.RowIndex >= 0)
            {
                vt = e;
                DataGridViewRow row = dgvVatTu.Rows[e.RowIndex];
                txtTimKiem.Text = sDK;
                txtTenPhong.Text = row.Cells[0].Value.ToString(); ;
                txtMaVatTu.Text = row.Cells[1].Value.ToString();
                txtTenVatTu.Text = row.Cells[2].Value.ToString();
                txtSoLuong.Text = row.Cells[3].Value.ToString();
                txtGhiChu.Text = row.Cells[4].Value.ToString();
            }
        }

        private void btThem_Click(object sender, EventArgs e)
        {
            ktThem = true;
            XoaTrang(); KhoaMo(false);
            txtMaVatTu.Focus();
        }

        private void btSua_Click(object sender, EventArgs e)
        {
            ktThem = false;
            KhoaMo(false);
            macu = txtMaVatTu.Text;
            txtTenVatTu.Focus();
            txtMaVatTu.Enabled = false;
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (txtMaVatTu.Text == "") return;
            if (MessageBox.Show($"ban muon xoa vat tu {txtTenVatTu.Text} ?", "THONG BAO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var objXoa = dbcontext.tbl_VatTus.Where(x => x.MaVatTu == txtMaVatTu.Text).SingleOrDefault();
                dbcontext.tbl_VatTus.DeleteOnSubmit(objXoa);
                dbcontext.SubmitChanges();
                XoaTrang();
                LayNguonVatTu();
            }
        }

        private void btGhi_Click(object sender, EventArgs e)
        {
            if (txtTenVatTu.Text == "")
            {
                MessageBox.Show("chua nhap ten vat tu", "THONG BAO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenVatTu.Focus();
                return;
            }

            if (ktThem == true)
            {
                if (dbcontext.tbl_VatTus.FirstOrDefault(x => x.MaVatTu == txtMaVatTu.Text) != null)
                {
                    MessageBox.Show("Vat Tu Nay Da Ton Tai. Vui Long Kiem Tra Ma Vat Tu", "THONG BAO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtMaVatTu.Focus();
                    return;
                }
                tbl_VatTu objThem = new tbl_VatTu();
                objThem.MaVatTu = txtMaVatTu.Text;
                objThem.TenVatTu = txtTenVatTu.Text;
                objThem.MaPhong = txtTenPhong.Text;
                objThem.SoLuong = int.Parse(txtSoLuong.Text);
                objThem.GhiChu = "";
                dbcontext.tbl_VatTus.InsertOnSubmit(objThem);
            }
            else
            {
                var objSua = dbcontext.tbl_VatTus.Where(x => x.MaVatTu == macu).SingleOrDefault();
                objSua.MaVatTu = txtMaVatTu.Text;
                objSua.TenVatTu = txtTenVatTu.Text;
                objSua.MaPhong = txtTenPhong.Text;
                objSua.SoLuong = int.Parse(txtSoLuong.Text);
                objSua.GhiChu = txtGhiChu.Text;
                //addfix
                    //var objXoa = dbcontext.tbl_VatTus.Where(x => x.MaVatTu == macu).SingleOrDefault();
                    //dbcontext.tbl_VatTus.DeleteOnSubmit(objXoa);
                    //dbcontext.SubmitChanges();
                    //LayNguonVatTu();
                //addfix
                dbcontext.tbl_VatTus.InsertOnSubmit(objSua);
                
            }
            dbcontext.SubmitChanges();
            KhoaMo(true);
            LayNguonVatTu();
            txtMaVatTu.Enabled = true;
        }

        private void btKhongGhi_Click(object sender, EventArgs e)
        {
            try
            {
                XoaTrang();
                KhoaMo(true);
                dgvVatTu_CellMouseClick(sender, vt);
                txtMaVatTu.Enabled = true;
            }
            catch (Exception ex) { }
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
            Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "E1");

            head.MergeCells = true;

            head.Value2 = title;

            head.Font.Bold = true;

            head.Font.Name = "Times New Roman";

            head.Font.Size = "20";

            head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // Tạo tiêu đề cột 

            Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A3", "A3");

            cl1.Value2 = "Tên Phòng";

            cl1.ColumnWidth = 12;

            Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B3", "B3");

            cl2.Value2 = "Mã Vật Tư";

            cl2.ColumnWidth = 12;

            Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C3", "C3");

            cl3.Value2 = "Tên Vật Tư";
            cl3.ColumnWidth = 20.0;

            Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");

            cl4.Value2 = "Số Lượng";

            cl4.ColumnWidth = 10.5;

            Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");

            cl5.Value2 = "Ghi Chú";

            cl5.ColumnWidth = 25.5;


            Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A3", "E3");

            rowHead.Font.Bold = true;

            // Kẻ viền

            rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            // Thiết lập màu nền

            rowHead.Interior.ColorIndex = 6;

            rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // Tạo mảng theo datatable

            object[,] arr = new object[dataTable.Rows.Count, dataTable.Columns.Count];

            //Chuyển dữ liệu từ DataTable vào mảng đối tượng

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                DataRow dataRow = dataTable.Rows[row];

                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    //if (dataRow[col].ToString().ToLower() == "null")
                    //{
                    //    arr[col] = null;
                    //}
                    arr[row, col] = dataRow[col];
                }
            }


            //Thiết lập vùng điền dữ liệu

            int rowStart = 4;

            int columnStart = 1;

            int rowEnd = rowStart + dataTable.Rows.Count -1;

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

            DataColumn col1 = new DataColumn("Ten Phong");
            DataColumn col2 = new DataColumn("Ma Vat Tu");
            DataColumn col3 = new DataColumn("Ten Vat Tu");
            DataColumn col4 = new DataColumn("So Luong");
            DataColumn col5 = new DataColumn("Ghi Chu");

            dataTable.Columns.Add(col1);
            dataTable.Columns.Add(col2);
            dataTable.Columns.Add(col3);
            dataTable.Columns.Add(col4);
            dataTable.Columns.Add(col5);

            foreach(DataGridViewRow dgvrow in dgvVatTu.Rows)
            {
                DataRow dtrow = dataTable.NewRow();

                dtrow[0] = dgvrow.Cells[0].Value;
                dtrow[1] = dgvrow.Cells[1].Value;
                dtrow[2] = dgvrow.Cells[2].Value;
                dtrow[3] = dgvrow.Cells[3].Value;
                dtrow[4] = dgvrow.Cells[4].Value;

                dataTable.Rows.Add(dtrow);
            }

            ExportFile(dataTable, "vattu", "Danh Sách Vật Tư");
        }
    }
}
