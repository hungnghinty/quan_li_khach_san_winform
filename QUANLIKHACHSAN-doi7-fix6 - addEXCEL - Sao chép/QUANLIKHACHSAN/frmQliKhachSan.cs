using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QUANLIKHACHSAN;


namespace QUANLIKHACHSAN
{
    public partial class frmQliKhachSan : Form
    {
        public frmQliKhachSan()
        {
            InitializeComponent();
        }
        
        private void dịchVụToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmDichVu frmDV = new frmDichVu();
            frmDV.MdiParent = this;
            frmDV.Show();
        }

        private void ketThucToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void mnuInHoaDon_Click(object sender, EventArgs e)
        {
            frmtblHoaDon frmHD = new frmtblHoaDon();
            frmHD.MdiParent = this;
            frmHD.Show();
        }


        private void mnuVatTu_Click_1(object sender, EventArgs e)
        {
            frmVatTu frmVT = new frmVatTu();
            frmVT.MdiParent = this;
            frmVT.Show();
        }

        private void quảnLíTàiKhoanToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
