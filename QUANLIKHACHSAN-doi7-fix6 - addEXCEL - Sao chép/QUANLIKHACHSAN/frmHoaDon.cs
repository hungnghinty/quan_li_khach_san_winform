using System;
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
    public partial class frmtblHoaDon : Form
    {
        DBQuanLiKhachSanDataContext dbcontext = new DBQuanLiKhachSanDataContext();
        string sDK = "";
        DataGridViewCellMouseEventArgs vtHopDong, vttblHoaDon;
        bool ktKetThucHopDong, ktThem;
        int idHoaDon ;
        int vtDichVu;
        ListViewItem itemLV;
        public frmtblHoaDon()
        {
            InitializeComponent();
        }
        public void LayNguonHopDong()
        {
            var dl = (from hd in dbcontext.tbl_HopDongs
                      join ph in dbcontext.tbl_Phongs on hd.MaPhong equals ph.MaPhong
                      where ((ckbHienTatCa.Checked == true? true: hd.DaKetThuc == false )
                      && (sDK == "" ? true : hd.MaHopDong.Contains(sDK) ))
                      select new
                      {
                          hd.MaHopDong,
                          hd.TenKhachHang,
                          ph.MaPhong,
                          hd.GIaThue,
                          hd.NgayVao,
                          hd.NgayRa,
                          hd.DaKetThuc
                      }
                   ).ToList();
            dgvHopDong.DataSource = dl;
            XoaTrang();
        }
        public void LayNguontblHoaDon()
        {
            var dl = dbcontext.tbl_HoaDons.Where(p => p.MaHopDong.Equals(txtSoHD.Text)).Select(p => new
            {
                p.MaHD,
                p.MaHopDong,
                p.NgayLap,
                p.TienGIam,
                p.TienPhat,
                p.SoTienTra,
                p.TongTien,
                p.GhiChu
            }).ToList();
            dgvHoaDon.DataSource = dl;
        }
        public void LayNguonCTtblHoaDon()
        {
            var dl = (from dv in dbcontext.tbl_DichVus
                      join ct in dbcontext.tbl_CTHoaDons on dv.MaDV equals ct.MaDV
                      where ct.MaHD == idHoaDon
                      select new
                      {
                          dv.MaDV,
                          dv.TenDV,
                          ct.GiaTIen,
                          ct.SoLuong
                      }).ToList();
            lvwDichVu.Items.Clear();
            for (int i = 0; i < dl.Count; i++)
            {
                itemLV = new ListViewItem(new[] { dl[i].MaDV.ToString(), dl[i].TenDV.ToString(),
                                            dl[i].GiaTIen.ToString(),dl[i].SoLuong.ToString(),
                                            (dl[i].SoLuong*dl[i].GiaTIen).ToString()});
                lvwDichVu.Items.Add(itemLV);
            }
        }
        public void LayNguonDichVu()
        {
            var dl = dbcontext.tbl_DichVus.Select(p => new { p.MaDV, p.TenDV, p.SoTien }).ToList();
            lvwDichVu.Items.Clear();
            for (int i = 0; i < dl.Count; i++)
            {
                itemLV = new ListViewItem(new[] { dl[i].MaDV.ToString(), dl[i].TenDV.ToString(),
                                            dl[i].SoTien.ToString(),"0", "0"});
                lvwDichVu.Items.Add(itemLV);
            }
        }
        public void LayNguonBoSung()
        {
            var dsMaDV = dbcontext.tbl_CTHoaDons.Where(p => p.MaHD == idHoaDon).Select(p => p.MaDV).ToList();
            var dl = dbcontext.tbl_DichVus.Where(p =>  !dsMaDV.Contains(p.MaDV)).ToList();
            for (int i = 0; i < dl.Count; i++)
            {
                itemLV = new ListViewItem(new[] { dl[i].MaDV.ToString(), dl[i].TenDV.ToString(),
                                            dl[i].SoTien.ToString(),"0",
                                            "0"});
                lvwDichVu.Items.Add(itemLV);
            }
        }
        private void frmtblHoaDon_Load(object sender, EventArgs e)
        {
            KhoaMo(true);
            LayNguonHopDong();
        }
        public void KhoaMo(bool b)
        {
            txtTimKiem.ReadOnly = !b; ckbHienTatCa.Enabled = b;
            dgvHoaDon.Enabled = b; dgvHopDong.Enabled = b;
            lvwDichVu.Enabled = !b; dtNgayLap.Enabled = !b;
            txtTienGiam.ReadOnly = b; txtTienPhat.ReadOnly = b;
            txtTienThu.ReadOnly = b; txtTienTra.ReadOnly = b;
            txtTinhTrang.Enabled = !b; btTinhTien.Enabled = !b;
            btThem.Enabled = b; btSua.Enabled = b;
            btXoa.Enabled = b; btKetThuc.Enabled = b;
            btGhi.Enabled = !b; btKhongGhi.Enabled = !b;
            btXoaSoLuong.Enabled = !b;
        }

        
        public void XoaTrang()
        {
            txtTienGiam.Text = "0"; txtTienPhat.Text = "0";
            txtTienThu.Text = "0"; txtTienTra.Text = "0";
            txtGhiChu.Text = "";
            lvwDichVu.Items.Clear();
        }

        private void lvwDichVu_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                sDK = txtTimKiem.Text;
                LayNguonHopDong();
            }
        }

        private void ckbHienTatCa_CheckedChanged(object sender, EventArgs e)
        {
            sDK = txtTimKiem.Text;
            LayNguonHopDong();
        }


        private void btXoaSoLuong_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lvwDichVu.Items.Count; i++)
            {
                lvwDichVu.Items[i].SubItems[3].Text = "0";
                lvwDichVu.Items[i].SubItems[4].Text = "0";
            }
        }

        private void btTinhTien_Click(object sender, EventArgs e)
        {
            if (txtTienGiam.Text == "") txtTienGiam.Text = "0";
            if (txtTienPhat.Text == "") txtTienPhat.Text = "0";
            if (txtTienTra.Text == "") txtTienTra.Text = "0";
            if (txtTienThu.Text == "") txtTienThu.Text = "0";
            double tienphong = 0, tiendv = 0, tiengiam = 0, tienphat = 0, sotien = 0, sl, dg;
            tienphong = double.Parse(txtGiaThuePhong.Text);
            tiengiam = double.Parse(txtTienGiam.Text);
            tienphat = double.Parse(txtTienPhat.Text);
            for (int i = 0; i < lvwDichVu.Items.Count; i++)
            {
                sl = double.Parse(lvwDichVu.Items[i].SubItems[3].Text);
                dg = double.Parse(lvwDichVu.Items[i].SubItems[2].Text);
                lvwDichVu.Items[i].SubItems[4].Text = (sl * dg).ToString();
                tiendv = tiendv + (sl * dg);
            }
            sotien = tienphong + tienphat - tiengiam + tiendv;
            txtTienThu.Text = sotien.ToString();
        }

        private void lvwDichVu_MouseUp(object sender, MouseEventArgs e)
        {
            ListView.SelectedListViewItemCollection itemSL = lvwDichVu.SelectedItems;
            if (itemSL.Count > 0)
            {
                ListViewHitTestInfo i = lvwDichVu.HitTest(e.X, e.Y);
                int cellTop = lvwDichVu.Top + i.SubItem.Bounds.Top;
                txtNhapSoLuong.Location = new Point(txtNhapSoLuong.Left, cellTop);
                txtNhapSoLuong.Text = itemSL[0].SubItems[3].Text;
                vtDichVu = itemSL[0].Index;
                txtNhapSoLuong.Visible = true;
                txtNhapSoLuong.Focus();

            }
        }

        private void txtNhapSoLuong_Leave(object sender, EventArgs e)
        {
            lvwDichVu.Items[vtDichVu].SubItems[3].Text = txtNhapSoLuong.Text;
            txtNhapSoLuong.Visible = false;
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            if (ktKetThucHopDong == true)
            {
                MessageBox.Show("Hop dong da ke thuc", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtSoHD.Text == "") return;
            if (idHoaDon == -1) return;
            if (MessageBox.Show("Ban co muon xoa hoa don nay?", "thong bao", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //addfix
                var cthd = dbcontext.tbl_CTHoaDons.Where(p => p.MaHD == idHoaDon).ToList();
                dbcontext.tbl_CTHoaDons.DeleteAllOnSubmit(cthd);
                //addfix
                tbl_HoaDon hd = dbcontext.tbl_HoaDons.Where(p => p.MaHD == idHoaDon).FirstOrDefault();
                dbcontext.tbl_HoaDons.DeleteOnSubmit(hd);
                dbcontext.SubmitChanges();
                XoaTrang();
                idHoaDon = -1;
                LayNguontblHoaDon();
            }
        }

        private void btSua_Click(object sender, EventArgs e)
        {
            if (ktKetThucHopDong == true)
            {
                MessageBox.Show("Hop dong da ke thuc", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtSoHD.Text == "") return;
            if (idHoaDon == -1) return;
            ktThem = false;
            KhoaMo(false);
            LayNguonBoSung();
        }

        private void btGhi_Click(object sender, EventArgs e)
        {
            btTinhTien_Click(sender, e);
            if (MessageBox.Show("ban co muon ghi hoa don khong?", "thong bao", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
            if (ktThem == true)
            {
                tbl_HoaDon hd = new tbl_HoaDon();
                //addfix
                hd.MaHD = dbcontext.tbl_HoaDons.Where(p => p.MaHopDong == txtSoHD.Text).Count();
                //addfix
                hd.MaHopDong = txtSoHD.Text;
                hd.NgayLap = dtNgayLap.Value;
                hd.TienGIam = int.Parse(txtTienGiam.Text);
                hd.TienPhat = int.Parse(txtTienPhat.Text);
                hd.SoTienTra = int.Parse(txtTienTra.Text);
                hd.TongTien = int.Parse(txtTienThu.Text);
                hd.GhiChu = txtGhiChu.Text;

                dbcontext.tbl_HoaDons.InsertOnSubmit(hd);
                dbcontext.SubmitChanges();

                tbl_HoaDon dl = dbcontext.tbl_HoaDons.OrderByDescending(p => p.MaHD).FirstOrDefault();
                idHoaDon = dl.MaHD;
            }
            else
            {
                tbl_HoaDon hd = dbcontext.tbl_HoaDons.Where(p => p.MaHD == idHoaDon).FirstOrDefault();
                //hd.MaHD = int.Parse(txtSoHD.Text);
                hd.NgayLap = dtNgayLap.Value;
                hd.TienGIam = decimal.Parse(txtTienGiam.Text);
                hd.TienPhat = decimal.Parse(txtTienPhat.Text);
                hd.SoTienTra = decimal.Parse(txtTienTra.Text);
                hd.TongTien = decimal.Parse(txtTienThu.Text);
                hd.GhiChu = txtGhiChu.Text;

                dbcontext.SubmitChanges();

                var dsxoa = dbcontext.tbl_CTHoaDons.Where(p => p.MaHD == idHoaDon).ToList();
                dbcontext.tbl_CTHoaDons.DeleteAllOnSubmit(dsxoa);
                dbcontext.SubmitChanges();
            }

            int sl; decimal dg;
            for (int i = 0; i < lvwDichVu.Items.Count; i++)
            {
                sl = int.Parse(lvwDichVu.Items[i].SubItems[3].Text);
                dg = decimal.Parse(lvwDichVu.Items[i].SubItems[2].Text);
                if (sl > 0)
                {
                    tbl_CTHoaDon ct = new tbl_CTHoaDon();
                    ct.MaHD = idHoaDon;
                    ct.MaDV = lvwDichVu.Items[i].Text;
                    ct.SoLuong = sl; ct.GiaTIen = dg;
                    dbcontext.tbl_CTHoaDons.InsertOnSubmit(ct);
                    dbcontext.SubmitChanges();
                }
            }
            KhoaMo(true);
            LayNguontblHoaDon();
            LayNguonCTtblHoaDon();
        }

        private void btKetThuc_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btThem_Click(object sender, EventArgs e)
        {
            if (ktKetThucHopDong == true)
            {
                MessageBox.Show("Hop dong da ke thuc", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtSoHD.Text == "") return;
            ktThem = true;
            KhoaMo(false);
            XoaTrang();
            LayNguonDichVu();
            dtNgayLap.Value = DateTime.Now;
            dtNgayLap.Focus();
        }

        private void dgvHoaDon_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvHoaDon.RowCount <= 0) return;
            if (e.RowIndex >= 0)
            {
                vttblHoaDon = e;
                DataGridViewRow row = dgvHoaDon.Rows[e.RowIndex];
                dtNgayLap.Value = DateTime.Parse(row.Cells[2].Value.ToString());
                txtTienGiam.Text = row.Cells[3].Value.ToString();
                txtTienPhat.Text = row.Cells[4].Value.ToString();
                txtTienTra.Text = row.Cells[5].Value.ToString();
                txtTienThu.Text = row.Cells[6].Value.ToString();
                //txtTinhTrang.Text = row.Cells[0].Value.ToString();
                txtGhiChu.Text = row.Cells[7].Value.ToString();

                txtMaHD.Text = row.Cells[0].Value.ToString();

                idHoaDon = int.Parse(row.Cells[0].Value.ToString());
                LayNguonCTtblHoaDon();
            }
        }

        private void dgvHopDong_CellMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvHopDong.Rows.Count <= 0) return;
            if (e.RowIndex >= 0)
            {
                vtHopDong = e;
                DataGridViewRow row = dgvHopDong.Rows[e.RowIndex];
                txtSoHD.Text = row.Cells[0].Value.ToString();
                txtKhachHang.Text = row.Cells[1].Value.ToString();
                txtSoPhong.Text = row.Cells[2].Value.ToString();
                txtGiaThuePhong.Text = row.Cells[3].Value.ToString();
                txtTuNgay.Text = row.Cells[4].Value.ToString();
                txtDenNgay.Text = row.Cells[5].Value.ToString();
                ktKetThucHopDong = (bool)row.Cells[6].Value;
                if (ktKetThucHopDong == true) txtTinhTrang.Text = "Đã Kết Thúc";
                else txtTinhTrang.Text = "Chưa Kết Thúc";
                XoaTrang();
                LayNguontblHoaDon();
            }
        }
        private void btKhongGhi_Click(object sender, EventArgs e)
        {
            try
            {
                KhoaMo(true);
                XoaTrang();
                dgvHoaDon_CellMouseClick(sender, vttblHoaDon);
            }
            catch (Exception ex) { }
        }


        private void btXuatExcel_Click_1(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();

            DataColumn col1 = new DataColumn("Ma Hoa Don");
            DataColumn col2 = new DataColumn("So Hop Dong");
            DataColumn col3 = new DataColumn("Ngay Lap");
            DataColumn col4 = new DataColumn("Tien Phat");
            DataColumn col5 = new DataColumn("Tien Giam");
            DataColumn col6 = new DataColumn("Khach Tra");
            DataColumn col7 = new DataColumn("Tong Tien");
            DataColumn col8 = new DataColumn("Ghi Chu");


            dataTable.Columns.Add(col1);
            dataTable.Columns.Add(col2);
            dataTable.Columns.Add(col3);
            dataTable.Columns.Add(col4);
            dataTable.Columns.Add(col5);
            dataTable.Columns.Add(col6);
            dataTable.Columns.Add(col7);
            dataTable.Columns.Add(col8);


            foreach (DataGridViewRow dgvrow in dgvHoaDon.Rows)
            {
                DataRow dtrow = dataTable.NewRow();

                dtrow[0] = dgvrow.Cells[0].Value;
                dtrow[1] = dgvrow.Cells[1].Value;
                dtrow[2] = dgvrow.Cells[2].Value;
                dtrow[3] = dgvrow.Cells[3].Value;
                dtrow[4] = dgvrow.Cells[4].Value;
                dtrow[5] = dgvrow.Cells[5].Value;
                dtrow[6] = dgvrow.Cells[6].Value;
                dtrow[7] = dgvrow.Cells[7].Value;

                dataTable.Rows.Add(dtrow);
            }

            ExportFileHoaDon(dataTable, "hoadon", "Danh Sách Hoá Đơn");
        }
        public void ExportFileHoaDon(DataTable dataTable, string sheetName, string title)
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
            Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "H1");

            head.MergeCells = true;

            head.Value2 = title;

            head.Font.Bold = true;

            head.Font.Name = "Times New Roman";

            head.Font.Size = "20";

            head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //thong tin hop dong
            Microsoft.Office.Interop.Excel.Range clKhachHang = oSheet.get_Range("C2", "C2");

            clKhachHang.Value2 = "Khách Hàng:" + txtKhachHang.Text;
            Microsoft.Office.Interop.Excel.Range clPhong = oSheet.get_Range("D2", "D2");

            clPhong.Value2 = "Phòng:" + txtSoPhong.Text;
            Microsoft.Office.Interop.Excel.Range clgia = oSheet.get_Range("E2", "E2");

            clgia.Value2 = "Giá Phòng:" + txtGiaThuePhong.Text;
            Microsoft.Office.Interop.Excel.Range clTuNgay = oSheet.get_Range("F2", "F2");

            clTuNgay.Value2 = "Từ Ngày:" + txtTuNgay.Text;
            Microsoft.Office.Interop.Excel.Range clDenNgay = oSheet.get_Range("G2", "G2");

            clDenNgay.Value2 = "Đến Ngày:" + txtDenNgay.Text;

            // Tạo tiêu đề cột 

            Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A3", "A3");

            cl1.Value2 = "Mã Hoá Đơn";

            cl1.ColumnWidth = 10.0;

            Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B3", "B3");

            cl2.Value2 = "Số Hợp Đồng";

            cl2.ColumnWidth = 10.0;

            Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C3", "C3");

            cl3.Value2 = "Ngày Lập";
            cl3.ColumnWidth = 25.0;

            Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");

            cl4.Value2 = "Tiền Phạt";

            cl4.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");

            cl5.Value2 = "Tiền Giảm";

            cl5.ColumnWidth = 20.0;
            Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");

            cl6.Value2 = "Khách Trả";

            cl6.ColumnWidth = 30.0;
            Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");

            cl7.Value2 = "Tổng Tiền";

            cl7.ColumnWidth = 30.0;
            Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");

            cl8.Value2 = "Ghi Chu";

            cl8.ColumnWidth = 25.5;


            Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A3", "H3");

            rowHead.Font.Bold = true;

            // Kẻ viền

            rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            // Thiết lập màu nền

            rowHead.Interior.ColorIndex = 7;

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
    }
}
