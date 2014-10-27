using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EStudio
{
    public partial class FormQr : Form
    {
        public object sQr = null;
        public FormQr()
        {
            InitializeComponent();
        }

        private void FormQr_Load(object sender, EventArgs e)
        {
            pictureBoxQR.Image = (Image)sQr;
        }

        private void btnQrExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
