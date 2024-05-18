using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QUANLIKHACHSAN
{
    internal class Public
    {
        public static SqlConnection conn;
        public static SqlCommand cmd;
        public static string sql = "";

        public static SqlConnection KetNoi()//using System.Data.SqlClient;
        {
            string connS = @"Data Source=.\SQLEXPRESS;Initial Catalog=DB_QLKS.bak;Integrated Security=True";
            conn = new SqlConnection(connS);
            return conn;
        }
        public static DataTable LayNguon(string sql)//using System.Data;
        {
            SqlDataAdapter da = new SqlDataAdapter(sql, KetNoi());
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public static void GanNguonDataGridView(DataGridView dgData, string sql)//using System.Windows.Forms;
        {
            dgData.DataSource = LayNguon(sql);
        }

        public static void ThucThiSQL(string sql)
        {
            cmd = new SqlCommand(sql, KetNoi());
            if (conn.State != ConnectionState.Open) conn.Open();
            cmd.ExecuteNonQuery();
        }

        public static void XoaDongDL(string sTenBang, string sTenTruong, string macu)
        {
            sql = "Delete From " + sTenBang + " Where " + sTenTruong + " = '" + macu + "'";
            cmd = new SqlCommand(sql, KetNoi());
            if (conn.State != ConnectionState.Open) conn.Open();
            cmd.ExecuteNonQuery();
        }
        // " + gt + "
        public static bool ktTrungMa(string sTenBang, string sTenTruong, bool ktThem, string mamoi, string macu)
        {
            sql = "select " + sTenTruong + " from " + sTenBang + " where " + sTenTruong + "='" + mamoi + "'";
            if (ktThem == false)
                sql = sql + " and " + sTenTruong + "<>'" + macu + "'";
            DataTable dt = LayNguon(sql);
            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }

    }
}
