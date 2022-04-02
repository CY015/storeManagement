using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FastReport;

namespace 订货系统
{
    public partial class Form1 : Form
    {
        private DataTable dt;
        public Form1(DataTable print)
        {
            InitializeComponent();
            this.dt = print.Copy();
            Report preport = new Report();
            preport.Load("F:\\1_University\\Project_C#\\reportTemplate\\bangdan.frx");

            preport.RegisterData(dt, "bangdan");
            preport.GetDataSource("bangdan").Enabled = true;

            preport.Show();
        }
    }
}
