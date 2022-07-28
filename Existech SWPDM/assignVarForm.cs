using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Existech_SWPDM
{
    public partial class assignVarForm : Form
    {
        // public variable for selected file type
        public string SelectedVariable = "";

        public assignVarForm()
        {
            InitializeComponent();
            confirmBtn.DialogResult = DialogResult.OK;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            SelectedVariable = "";
        }

        private void confirmBtn_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                SelectedVariable = "94";
            if (radioButton2.Checked)
                SelectedVariable = "96";
            Close();
        }
    }
}
