using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QUANLIKHACHSAN
{
    public partial class frmDangNhap : Form
    {
        #region Variables
        string ConnectionString = ConfigurationManager.AppSettings["ConnectionString"];
        SqlConnection conn;
        DataSet ds;
        DBQuanLiKhachSanDataContext dbcontext = new DBQuanLiKhachSanDataContext();
        // sao ko choi linq het di
        // cai kia cop sang ma . doi linq nhieu loi lam
        // the h sao :>
        // ko dc thi thoi =))
        // doi 
        //det

        #endregion

        public frmDangNhap()
        {
            InitializeComponent();
        }
        private bool CheckEmpty(TextBox obj, string Message)
        {
            if (obj.Text.Trim() == "")
            {
                MessageBox.Show(Message, "Thông báo");
                obj.Focus();
                return false;
            }
            return true;
        }


        private void FrmDangNhap_Load(object sender, EventArgs e)
        {

        }


        private void btnDangnhap_Click(object sender, EventArgs e)
        {
            try
            {

                // khai báo SqlConnection
                conn = new SqlConnection(ConnectionString);
                if (CheckEmpty(txtUserName, "Chưa nhập thông tin tài khoản") == false)
                {
                    return;
                }
                if (CheckEmpty(txtPassword, "Chưa nhập thông tin mật khẩu") == false)
                {
                    return;
                }
                this.Hide();
                frmQliKhachSan frmain = new frmQliKhachSan();
                frmain.ShowDialog();
                this.Close();
                //    //Mở kết nối 
                //    if (conn.State == ConnectionState.Closed)
                //        conn.Open();
                //    //Khai báo câu truy vấn 
                //    string sql = "Select * from tbl_TaiKhoan where TenTaiKhoan = '" + txtUserName.Text + "' And MatKhau = '" + txtPassword.Text + "' ";

                //    //Khai báo SqlCommand, SqlDataAdapter
                //    SqlCommand command = new SqlCommand(sql, conn);

                //    SqlDataAdapter da = new SqlDataAdapter(command);
                //    //Khai báo Dataset
                //    ds = new DataSet();
                //    //fill dữ liệu 
                //    da.Fill(ds);

                //    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                //    {

                //        //MessageBox.Show("Đăng nhập thành công ", "Thông báo    


                //    }
                //    else
                //    {
                //        MessageBox.Show("Thông tin tài khoản hoặc mật khẩu chưa chính xác ", "Thông báo");
                //    }
            }

            catch (Exception ex)
            {

                MessageBox.Show("Lỗi: " + ex.Message, "Thông báo");
            }
            finally
            {
                conn.Dispose();
            }

        }

        private void label4_Click(object sender, EventArgs e)
        {
            if (txtPassword.PasswordChar == '*')
            {
                label4.Text = "hide";
                txtPassword.PasswordChar = '\0';
            }
            else
            {
                txtPassword.PasswordChar = '*';
                label4.Text = "show";
            }
        }
    }
}
