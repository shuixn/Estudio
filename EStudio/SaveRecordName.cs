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
    public partial class SaveRecordName : Form
    {
        public SaveRecordName()
        {
            InitializeComponent();
        }

        private void SaveRecordName_Load(object sender, EventArgs e)
        {

        }
        public string saveName;

        private void btnSave_Click(object sender, EventArgs e)
        {
                saveName = this.tbxRecordName.Text.Trim();
                this.Close();
        }

    }
}
