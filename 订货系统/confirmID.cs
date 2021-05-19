using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 订货系统
{
    public partial class confirmID : Form
    {
        public confirmID()
        {
            InitializeComponent();
        }
        public class ID
        {
            public static int identity;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string password = tBpwd.Text.Trim();
            if (this.cBID.SelectedIndex == 0) 
            {
                ID.identity = 0;    //管理员
                if (password == "1")
                {
                    this.DialogResult = DialogResult.OK;
                }
                else if(password == "")
                {
                    MessageBox.Show("请输入密码", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                    MessageBox.Show("密码错误！请重新输入", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (this.cBID.SelectedIndex == 1)
            {
                ID.identity = 1;    //订货员
                this.DialogResult = DialogResult.OK;
            }
            else
                MessageBox.Show("请选择登陆身份","",MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

    }
}
