using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using NPOI;
using FastReport;

namespace 订货系统
{
    public partial class mainForm : Form
    {
        DataTable ioInfo = new DataTable();
        public mainForm()
        {
            InitializeComponent();

            /*************订货员权限*************/
            /***********只能看订货报表***********/
            if (confirmID.ID.identity == 1)
            {
                TabPage tp0 = tabControl1.TabPages[0];
                TabPage tp1 = tabControl1.TabPages[1];
                TabPage tp2 = tabControl1.TabPages[2];
                TabPage tp3 = tabControl1.TabPages[3];
                TabPage tp4 = tabControl1.TabPages[4];
                TabPage tp5 = tabControl1.TabPages[5];

                tabControl1.TabPages.Remove(tp0);
                tabControl1.TabPages.Remove(tp1);
                tabControl1.TabPages.Remove(tp2);
                tabControl1.TabPages.Remove(tp3);
                tabControl1.TabPages.Remove(tp4);
                tabControl1.TabPages.Remove(tp5);
                btnOrderReport2.Visible = false;
            }
        }

        /// <summary>
        /// 事务功能选择
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /***********更新零件信息&查询库存信息************/
            if (this.comboBox1.SelectedIndex == 0 || this.comboBox1.SelectedIndex == 4)
            {
                this.tabControl1.SelectedIndex = 1;
            }
            /***********更新供应商信息************/
            else if (this.comboBox1.SelectedIndex == 1)
            {
                this.tabControl1.SelectedIndex = 2;
            }
            /***********出入库************/
            else if (this.comboBox1.SelectedIndex == 2)
            {
                this.tabControl1.SelectedIndex = 3;
            }
            /***********订购零件************/
            else if (this.comboBox1.SelectedIndex == 3)
            {
                this.tabControl1.SelectedIndex = 6;
            }
            /***********查询往日订货信息************/
            else if (this.comboBox1.SelectedIndex == 5)
            {
                this.tabControl1.SelectedIndex = 5;
            }
        }

        /// <summary>
        /// 返回&退出
        /// </summary>
        private void btnReturn_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("确认退出", "Tips", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                Application.Exit();
        }
        private void btnReturn2_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }
        private void btnExit2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认退出", "Tips", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                Application.Exit();
        }

        /// <summary>
        /// 增删改查 零件数据
        /// </summary>
        private void btnComSearch_Click(object sender, EventArgs e)
        {
            string cID;
            string cName;
            string searchid;  //查找条件
            ConnSQL con = new ConnSQL();
            DataTable comInfo = new DataTable();
            cID = comID.Text.ToString().Trim();
            cName = comName.Text.ToString().Trim();

            //没有填写零件编号
            if (cID == "")
            {
                //有无填写零件名称
                if (cName == "")
                {
                    searchid = "SELECT * FROM Component";
                    comInfo = con.ExecuteQuery(searchid);
                    dGV_component.DataSource = comInfo;
                }
                else
                {
                    cName = cName + '%';
                    searchid = "SELECT * FROM Component WHERE cy_cName LIKE '" + cName + "'";
                    comInfo = con.ExecuteQuery(searchid);
                    dGV_component.DataSource = comInfo;
                }
            }

            //填了零件编号
            else
            {
                cID = cID + '%';
                searchid = "SELECT * FROM Component WHERE cy_cID LIKE '" + cID + "'";
                comInfo = con.ExecuteQuery(searchid);
                dGV_component.DataSource = comInfo;
            }

            dGV_component.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_component.Columns["cy_cName"].HeaderText = "零件名称";
            dGV_component.Columns["cy_cType"].HeaderText = "零件型号";
            dGV_component.Columns["cy_cNum"].HeaderText = "库存数量";
            dGV_component.Columns["cy_cMinNum"].HeaderText = "库存临界值";
            dGV_component.Columns["cy_cPrice"].HeaderText = "目前单价";
            dGV_component.Columns["cy_sID"].HeaderText = "供应商编号";

            if (dGV_component == null || dGV_component.Rows.Count <= 0)
            {
                MessageBox.Show("库存中无零件数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnComUpdate_Click(object sender, EventArgs e)
        {
            Component c1 = new Component();     //零件
            string cID = this.comID.Text.ToString().Trim();
            string cName = this.comName.Text.ToString().Trim();
            string cType = this.comType.Text.ToString().Trim();
            string comsID = csID.Text.ToString().Trim();
            float cPrice;
            int cLess;
            string seConditon;      //查找条件
            ConnSQL seCon = new ConnSQL();      //查找的连接实例
            DataTable updateInfo = new DataTable();

            seConditon = "SELECT * FROM Component WHERE cy_cID='" + cID + "'";
            updateInfo = seCon.ExecuteQuery(seConditon);
            dGV_component.DataSource = updateInfo;
            dGV_component.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_component.Columns["cy_cName"].HeaderText = "零件名称";
            dGV_component.Columns["cy_cType"].HeaderText = "零件型号";
            dGV_component.Columns["cy_cNum"].HeaderText = "库存数量";
            dGV_component.Columns["cy_cMinNum"].HeaderText = "库存临界值";
            dGV_component.Columns["cy_cPrice"].HeaderText = "目前单价";
            dGV_component.Columns["cy_sID"].HeaderText = "供应商编号";

            //判断价格的格式
            if (comPrice.Text != "")
            {
                string priceInput = comPrice.Text.ToString();
                Regex price = new Regex(@"^(([0-9]{1,3})+(.[0-9]{1,2})?)|([0-9]{1-3})$");   //0.00-999.99

                if (price.IsMatch(priceInput))
                    cPrice = float.Parse(priceInput);
                else
                {
                    cPrice = 0.0F;
                    MessageBox.Show("请正确填写零件价格。范围为0-999.99", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                cPrice = -0.1F;

            //判断临界值的格式
            if (comLess.Text != "")
            {
                string lessInput = comLess.Text.ToString();
                Regex less = new Regex(@"^[0-9]{1,3}$");  //0-999

                if (less.IsMatch(lessInput))
                    cLess = int.Parse(lessInput);
                else
                {
                    cLess = 0;
                    MessageBox.Show("请正确填写零件临界值。范围为0-999", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }        
            }
            else
                cLess = -1;

            //插入新信息
            if (dGV_component.Rows.Count == 0 && cID != "")
            {
/*                if(cPrice == 0.0F)
                {
                    MessageBox.Show("请正确填写零件价格。范围为0-999.99", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (cLess == 0)
                {
                    MessageBox.Show("请正确填写零件临界值。范围为0-999", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }*/
                if (MessageBox.Show("此零件将作为新零件: " + cID + " 加入到库存中", "入库", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    c1.insertInfo(cID, cName, cType, cPrice, cLess, comsID);
                    MessageBox.Show("添加成功！请按查询按钮查看最新零件信息。", "提示", MessageBoxButtons.OK);
                }
            }
            //更改信息
            else if(dGV_component.Rows.Count != 0 && cID != "")
            {
                c1.updateInfo(cID, cName, cType, cPrice, cLess, comsID);
                if(cPrice>0||cLess>0||cName!=""||cType!=""||comsID!="")
                    MessageBox.Show("零件信息更新成功！请按查询按钮查看最新零件信息。", "提示", MessageBoxButtons.OK);
            }
            else
            //MessageBox.Show("添加成功！请按查询按钮查看最新零件信息。", "提示", MessageBoxButtons.OK);
                MessageBox.Show("零件信息更新失败，请检查零件编号是否过长或是无填写", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void btnComDel_Click(object sender, EventArgs e)
        {
            string delCom;
            int i = this.dGV_component.Rows.Count;
            ConnSQL del = new ConnSQL();

            try
            {
                int iCount = dGV_component.SelectedRows.Count;
                if (iCount < 1)
                {
                    MessageBox.Show("未选定所要删除的零件所在行", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (MessageBox.Show("确认是否删除选中的零件信息？\n（此操作不可撤回！）", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    string cid;
                    if (i != this.dGV_component.Rows.Count - 2)
                    {
                        for (int k = i; k >= 1; k--)
                        {
                            //被选中状态则删除
                            if (this.dGV_component.Rows[k - 1].Selected == true)
                            {
                                cid = this.dGV_component.Rows[k - 1].Cells[0].Value.ToString();
                                dGV_component.Rows.Remove(dGV_component.Rows[k - 1]);

                                delCom = "delete from Component where cy_cID = '" + cid + "'";                                           
                                del.ExecuteUpdate(delCom);
                            }
                        }
                    }
                    //逆序删除防止串行
                    MessageBox.Show("删除成功");
                }
                else
                    MessageBox.Show("未成功删除所选零件");
            }
            catch (Exception)
            {
                MessageBox.Show("未成功删除");
            }
        }

        /// <summary>
        /// 增删改查 供应商数据
        /// </summary>
        /// TODO:删除
        private void btnSupSearch_Click(object sender, EventArgs e)
        {
            string sID = this.supID.Text.ToString().Trim();
            string searchid;  //查找条件
            ConnSQL con = new ConnSQL();
            DataTable supInfo = new DataTable();
            

            //没有填写供应商编号
            if (sID == "")
            {
                searchid = "SELECT * FROM Supply";
                supInfo = con.ExecuteQuery(searchid);
                dGV_Supply.DataSource = supInfo;
            }

            //填了供应商编号
            else
            {
                sID = sID +'%';
                searchid = "SELECT * FROM Supply WHERE cy_sID LIKE '" + sID + "'";
                supInfo = con.ExecuteQuery(searchid);
                dGV_Supply.DataSource = supInfo;
            }

            dGV_Supply.Columns["cy_sID"].HeaderText = "供应商编号";
            dGV_Supply.Columns["cy_sName"].HeaderText = "供应商名称";
            dGV_Supply.Columns["cy_sContact"].HeaderText = "联系方式";

            if (dGV_Supply == null || dGV_Supply.Rows.Count <= 0)
            {
                MessageBox.Show("无供应商数据，请添加", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnSupUpdate_Click(object sender, EventArgs e)
        {
            Supply s1 = new Supply();   //供应商
            string sID = this.supID.Text.ToString().Trim();
            string sName = this.supName.Text.ToString().Trim();
            string sCon = this.supCon.Text.ToString().Trim();
            string seCondition;      //查找条件
            ConnSQL seCon = new ConnSQL();      //查找的连接实例
            DataTable updateInfo = new DataTable();

            seCondition = "SELECT * FROM Supply WHERE cy_sID='" + sID + "'";
            updateInfo = seCon.ExecuteQuery(seCondition);
            dGV_Supply.DataSource = updateInfo;
            dGV_Supply.Columns["cy_sID"].HeaderText = "供应商编号";
            dGV_Supply.Columns["cy_sName"].HeaderText = "供应商名称";
            dGV_Supply.Columns["cy_sContact"].HeaderText = "联系方式";

            //联系方式格式是否正确
            if (supCon.Text != "")
            {
                string conInput = supCon.Text.ToString();
                Regex con = new Regex(@"^(([0-9]{3,4})+(-[1-9][0-9]{7}))|([1][0-9]{12})$");   // 固定电话 或者 11位电话号码

                if (con.IsMatch(conInput))
                    sCon = conInput;
                else
                {
                    sCon = "";
                    MessageBox.Show("联系方式输入格式有误，格式为三或四位的区号-八位固定电话号码xxx(xxxx)-xxxxxxxx或者11位移动电话号码", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
                sCon = "";


            //插入新信息
            if (dGV_Supply.Rows.Count == 0 && sID != "")
            {
                s1.insertInfo(sID, sName, sCon);
                MessageBox.Show("添加成功！请按查询按钮查看最新供应商信息。", "提示", MessageBoxButtons.OK);
            }
            //更改信息
            else if (dGV_Supply.Rows.Count != 0 && sID != "" && !(sName ==""&&sID ==""))
            {
                s1.updateInfo(sID, sName, sCon);
                MessageBox.Show("添加成功！请按查询按钮查看最新供应商信息。", "提示", MessageBoxButtons.OK);
            }
            else 
            MessageBox.Show("供应商信息更新失败" + "\n请检查供应商编号是否有填写", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
        private void btnSupDel_Click(object sender, EventArgs e)
        {
            string delSup;
            int i = this.dGV_Supply.Rows.Count;
            ConnSQL del = new ConnSQL();

            try
            {
                int iCount = dGV_Supply.SelectedRows.Count;
                if (iCount < 1)
                {
                    MessageBox.Show("未选定所要删除的供应商所在行", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (MessageBox.Show("确认是否删除选中的供应商信息？\n（此操作不可撤回！）", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    string sid;
                    if (i != this.dGV_Supply.Rows.Count - 2)
                    {
                        for (int k = i; k >= 1; k--)
                        {
                            //被选中状态则删除
                            if (this.dGV_Supply.Rows[k - 1].Selected == true)
                            {
                                sid = this.dGV_Supply.Rows[k - 1].Cells[0].Value.ToString();
                                dGV_Supply.Rows.Remove(dGV_Supply.Rows[k - 1]);

                                delSup= "DELETE FROM Supply WHERE cy_sID = '" + sid + "'";
                                del.ExecuteUpdate(delSup);
                            }
                        }
                    }
                    //逆序删除防止串行
                    MessageBox.Show("删除成功");
                }
                else
                    MessageBox.Show("未成功删除");
            }
            catch (Exception)
            {
                MessageBox.Show("未成功删除");
            }
        }

        /// <summary>
        /// 零件 出入库 查询
        /// </summary>
        private void btncIO_Click(object sender, EventArgs e)
        {

            //记录下ID和出入库数量更新至IO表
            //记录下出入库时间
            string cID = comIO_cID.Text.ToString();
            int cIONum = 0;
            string condition1;      //insert或者update 更新到零件表的
            string condition2;      //添加出入库记录信息到Component_IO表
            string Date = DateTime.Now.ToString("yyyy-MM-dd ");
            string Time = DateTime.Now.ToShortTimeString().ToString();
            string ioTime = Date + Time;
            ConnSQL ioCon = new ConnSQL();      //查找的连接实例
            DataTable updateInfo = new DataTable();     //零件清单
            Boolean newCom;     //标记是否为一个新零件加入库存中
            string orderTime = Date + Time;

            condition1 = "SELECT * FROM Component WHERE cy_cID='" + cID + "'";
            updateInfo  = ioCon.ExecuteQuery(condition1);
            dGV_component.DataSource = updateInfo;
            dGV_component.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_component.Columns["cy_cName"].HeaderText = "零件名称";
            dGV_component.Columns["cy_cType"].HeaderText = "零件型号";
            dGV_component.Columns["cy_cNum"].HeaderText = "库存数量";
            dGV_component.Columns["cy_cMinNum"].HeaderText = "库存临界值";
            dGV_component.Columns["cy_cPrice"].HeaderText = "目前单价";
            dGV_component.Columns["cy_sID"].HeaderText = "供应商编号";

            if (cID == "")
            {
                MessageBox.Show("请输入零件编号！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //判断有无数量&格式是否正确
            if (comIO_cNum.Text != "")
            {
                string input = comIO_cNum.Text.ToString();
                Regex inNum = new Regex(@"^[0-9]{1,3}$");  //0-999

                if (inNum.IsMatch(input))
                    cIONum = int.Parse(input);
                else
                {
                    MessageBox.Show("请正确输入零件出/入库数量！(数量范围:1-999)", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cIONum = -1;
                }
            }
            else
            {
                MessageBox.Show("请输入零件出/入库数量！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cIONum = -1;
            }

            //入库
            if (rBtnIN.Checked)
            {
                if(dGV_component.Rows.Count == 0) 
                {
                    if(cIONum > 0)      //操作数量大于0
                    {
                        if (MessageBox.Show("此零件将作为新零件: " + cID + " 加入到库存中", "入库", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            condition1 = "INSERT INTO Component(cy_cID,cy_cNum) VALUES('" + cID + "','" + cIONum + "')";
                            MessageBox.Show("零件入库成功！请在 零件信息 页面查询最新零件信息", "", MessageBoxButtons.OK);
                            newCom = true;
                        }
                        else
                        {
                            MessageBox.Show("零件入库失败！您取消了此次操作", "提示", MessageBoxButtons.OK);
                            newCom = false;
                        }
                    } 
                    else
                        newCom = false;
                }
                else
                {
                    if (cIONum > 0)
                    {
                        condition1 = "UPDATE Component SET cy_cNum = cy_cNum +'" + cIONum + "' WHERE cy_cID = '" + cID + "'";
                        MessageBox.Show("零件入库成功！请在 零件信息 页面查询最新零件信息", "", MessageBoxButtons.OK);
                        newCom = true;
                    }
                    else
                        newCom = false;
                }
                ioCon.ExecuteUpdate(condition1);

                if (newCom)
                {
                    condition2 = "INSERT INTO Component_IO(cy_cID,cy_cIOType,cy_cIONum,cy_cIOTime) VALUES('" + cID + "','入库','" + cIONum + "','" + ioTime + "')";
                    ioCon.ExecuteUpdate(condition2);
                }
            }

            //出库
            else if (rBtnOUT.Checked)
            {
                if(dGV_component.Rows.Count == 0)
                {
                    MessageBox.Show("库存中无此零件，请重新输入零件编号", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    try
                    {
                        if (cIONum > 0)
                        {
                            int minNum = int.Parse(dGV_component.Rows[0].Cells[4].Value.ToString());    //临界值
                            int kucun = int.Parse(dGV_component.Rows[0].Cells[3].Value.ToString());     //未出库前库存
                            int kucun2 = kucun - cIONum;     //出库后
                            if(kucun2<minNum)
                            {
                                MessageBox.Show("库存已小于临界值，出库失败，请添加此零件的订货信息。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                string condition = "INSERT INTO Component_Order(cy_cID,cy_sID,cy_orderNum,cy_orderTime) VALUES('" + cID + "','','"+kucun+"','" + orderTime + "')";
                                ioCon.ExecuteUpdate(condition);
                            }
                            else
                            {
                                condition1 = "UPDATE Component SET cy_cNum = cy_cNum -'" + cIONum + "' WHERE cy_cID = '" + cID + "'";
                                ioCon.ExecuteUpdate(condition1);
                                MessageBox.Show("零件出库成功！请在 零件信息 页面查询最新零件信息", "", MessageBoxButtons.OK);

                                condition2 = "INSERT INTO Component_IO(cy_cID,cy_cIOType,cy_cIONum,cy_cIOTime) VALUES('" + cID + "','出库','" + cIONum + "','" + ioTime + "')";
                                ioCon.ExecuteUpdate(condition2);
                            }
                        }
                        else
                            MessageBox.Show("零件出库失败！", "警告", MessageBoxButtons.OK,MessageBoxIcon.Error);

                    }
                    catch (Exception)
                    {
                        MessageBox.Show("出库数量大于零件库存现有量，出库失败", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            //没选任何功能
            else
                MessageBox.Show("请选择出/入库", "警告", MessageBoxButtons.OK,MessageBoxIcon.Warning);
        }
        private void cIO_Search_Click(object sender, EventArgs e)
        {
            string startTime = date_cIOStart.Value.ToString("yyyy-MM-dd");
            string endTime = date_cIOEnd.Value.ToString("yyyy-MM-dd") + "24:00:00";
            string condition;
            ConnSQL con = new ConnSQL();
            // DataTable ioInfo = new DataTable();

            condition = "SELECT cy_cID,cy_cIOType,cy_cIONum,cy_cIOTime FROM Component_IO WHERE cy_cIOTime BETWEEN '"+startTime+"' AND'"+endTime+ "' ORDER BY cy_cIOTime DESC";
            ioInfo = con.ExecuteQuery(condition);
            dGV_ComponentIO.DataSource = ioInfo;
            dGV_ComponentIO.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_ComponentIO.Columns["cy_cIOType"].HeaderText = "出/入库";
            dGV_ComponentIO.Columns["cy_cIONum"].HeaderText = "出/入库数量";
            dGV_ComponentIO.Columns["cy_cIOTime"].HeaderText = "出/入库时间";
        }
        private void btn_bangdan_Click(object sender, EventArgs e)
        {
            int i = this.dGV_ComponentIO.Rows.Count;
            DataTable dt_bangdan = new DataTable();
            try
            {
                int iCount = dGV_ComponentIO.SelectedRows.Count;
                if (iCount < 1)
                {
                    MessageBox.Show("无出入库详情，请先查询", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (iCount == 1)
                {
                    if (i != this.dGV_ComponentIO.Rows.Count - 2)
                    {
                        for (int k = i; k >= 1; k--)
                        {
                            //被选中状态
                            if (this.dGV_ComponentIO.Rows[k - 1].Selected == true)
                            {
                                dt_bangdan.Columns.Add("cID");
                                dt_bangdan.Columns.Add("cIOType");
                                dt_bangdan.Columns.Add("cIONum");
                                dt_bangdan.Columns.Add("cIOTime");
                                DataRow row = dt_bangdan.NewRow();

                                dt_bangdan.TableName = "bangdan";
                                row["cID"] = this.dGV_ComponentIO.Rows[k - 1].Cells[0].Value.ToString();
                                row["cIOType"] = this.dGV_ComponentIO.Rows[k - 1].Cells[1].Value.ToString();
                                row["cIONum"] = this.dGV_ComponentIO.Rows[k - 1].Cells[2].Value.ToString();
                                row["cIOTime"] = this.dGV_ComponentIO.Rows[k - 1].Cells[3].Value.ToString();

                                dt_bangdan.Rows.Add(row);
                                // dGV_component.Rows.Remove(dGV_component.Rows[k - 1]);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("行数超过一行");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("未成功打印");
            }

            Form1 f1 = new Form1(dt_bangdan);
        }

        /// <summary>
        /// 订货 刷新 导出到Excel 删除记录
        /// </summary>
        private void btnAdd_Click(object sender, EventArgs e)
        {
            //零件ID,供应商ID,订货数量添加到订货表,记录时间
            string orderCID = ocID.Text.ToString();
            string orderSID = osID.Text.ToString();
            int orderNum = 0;
            string Date = DateTime.Now.ToString("yyyy-MM-dd ");
            string Time = DateTime.Now.ToShortTimeString().ToString();
            string orderTime = Date + Time;
            string condition;
            string warningMessage = null;
            ConnSQL oCon = new ConnSQL();
            DataTable oList = new DataTable();

            if (ocID.Text == "")
                warningMessage = " 零件编号 ";
            if (osID.Text == "")
                warningMessage += "供应商编号 ";
            //检查零件数量格式是否正确
            if (ocNum.Text != "")
            {
                string input = ocNum.Text.ToString();
                Regex inNum = new Regex(@"^[0-9]{1,3}$");  //0-999

                if (inNum.IsMatch(input))
                    orderNum = int.Parse(input);
                else
                    MessageBox.Show("请正确输入订购零件数量！(数量范围:1-999)", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                warningMessage += "订购数量 ";

            if(orderCID==""|orderSID==""|orderNum==0)
                MessageBox.Show("请补充 "+warningMessage+"信息", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                condition = "SELECT cy_cID,cy_sID,cy_orderNum,cy_orderTime FROM Component_Order WHERE cy_cID = '" + orderCID + "'";
                dGV_Order.DataSource = oCon.ExecuteQuery(condition);
                if (dGV_Order.Rows.Count > 0)
                {
                    if (orderNum >= 0)
                    {
                        MessageBox.Show("将在原有的订货数据上增加订购零件数量");
                        condition = "UPDATE Component_Order SET cy_orderNum = cy_orderNum + '" + orderNum + "' WHERE cy_cID = '" + orderCID + "'";
                        oCon.ExecuteUpdate(condition);
                    }
                    if (orderSID != "")
                    {
                        MessageBox.Show("将在原有的订货数据上更改供应商编号");
                        condition = "UPDATE Component_Order SET cy_sID = '" + orderSID + "' WHERE cy_cID = '" + orderCID + "'";
                        oCon.ExecuteUpdate(condition);
                    }
                }
                else
                {
                    condition = "INSERT INTO Component_Order(cy_cID,cy_sID,cy_orderNum,cy_orderTime) VALUES('" + orderCID + "','" + orderSID + "','" + orderNum + "','" + orderTime + "')";
                    oCon.ExecuteUpdate(condition);
                }
                
            }

            condition = "SELECT cy_cID,cy_sID,cy_orderNum,cy_orderTime FROM Component_Order WHERE cy_cID = '" + orderCID + "'";
            dGV_Order.DataSource = oCon.ExecuteQuery(condition);

            //dGV_Order.Columns["cy_ID"].HeaderText = "订货序号";
            dGV_Order.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_Order.Columns["cy_sID"].HeaderText = "供应商编号";
            dGV_Order.Columns["cy_orderNum"].HeaderText = "订货数量";
            dGV_Order.Columns["cy_orderTime"].HeaderText = "订货时间";

            if (dGV_Order.Rows.Count == 1)
                MessageBox.Show("已添加零件 " + orderCID + "的订货信息"+"\n请点击刷新查看当日全部订货信息", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("添加失败", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btn_orderFresh_Click(object sender, EventArgs e)
        {
            string search;
            ConnSQL conn = new ConnSQL();
            string Date = DateTime.Now.ToString("yyyy-MM-dd")+'%';
            DataTable orderList = new DataTable();

            search = "SELECT cy_cID,cy_sID,cy_orderNum,cy_orderTime FROM Component_Order WHERE cy_orderTime LIKE '" + Date+ "' ORDER BY cy_orderTime DESC";
            dGV_Order.DataSource = conn.ExecuteQuery(search);

            //dGV_Order.Columns["cy_ID"].HeaderText = "订货序号";
            dGV_Order.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_Order.Columns["cy_sID"].HeaderText = "供应商编号";
            dGV_Order.Columns["cy_orderNum"].HeaderText = "订货数量";
            dGV_Order.Columns["cy_orderTime"].HeaderText = "订货时间";

            if (dGV_Order.Rows.Count == 0)
                MessageBox.Show("今日无订货数据", "警告",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            
        }
        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            if (dGV_Order.Rows.Count == 0)
            {
                MessageBox.Show("当前无数据导出！请先刷新今日订货信息，确认当前有无当日订货信息。");
            }
            else
            {
                string Date = DateTime.Now.ToString("yyyy-MM-dd");
                SaveFileDialog s1 = new SaveFileDialog();

                s1.Title = "请选择要导出的位置";
                s1.Filter = "Excel文件|*.xls";
                s1.FileName = Date + "订货信息";
                if (s1.ShowDialog() == DialogResult.OK)
                {
                    NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                    NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet(Date);
                    NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(0);

                    row1.CreateCell(0).SetCellValue("订货序号");
                    row1.CreateCell(1).SetCellValue("零件编号");
                    row1.CreateCell(2).SetCellValue("供应商编号");
                    row1.CreateCell(3).SetCellValue("订货数量");
                    row1.CreateCell(4).SetCellValue("订货时间");

                    for (int i = 0; i < dGV_Order.Rows.Count; i++)
                    {
                        NPOI.SS.UserModel.IRow rowTemp = sheet1.CreateRow(i + 1);
                        rowTemp.CreateCell(0).SetCellValue(i + 1);
                        rowTemp.CreateCell(1).SetCellValue(dGV_Order.Rows[i].Cells[0].Value.ToString());
                        rowTemp.CreateCell(2).SetCellValue(dGV_Order.Rows[i].Cells[1].Value.ToString());
                        rowTemp.CreateCell(3).SetCellValue(dGV_Order.Rows[i].Cells[2].Value.ToString());
                        rowTemp.CreateCell(4).SetCellValue(dGV_Order.Rows[i].Cells[3].Value.ToString());
                    }

                    FileStream orderRepo = File.OpenWrite(s1.FileName.ToString());
                    try
                    {
                        book.Write(orderRepo);
                        orderRepo.Seek(0, SeekOrigin.Begin);
                        MessageBox.Show("导出成功！", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("导出失败！", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        throw;
                    }
                    finally
                    {
                        if (orderRepo != null)
                        {
                            orderRepo.Close();
                        }
                    }
                }
            }
        }
        private void btnOrderDel_Click(object sender, EventArgs e)
        {
            string delOrder;
            int i = this.dGV_Order.Rows.Count;
            ConnSQL del = new ConnSQL();

            try
            {
                int iCount = dGV_Order.SelectedRows.Count;
                if (iCount < 1)
                {
                    MessageBox.Show("未选定所要删除的订货记录所在行", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (MessageBox.Show("确认是否删除选中的订货记录？\n（此操作不可撤回！！！）", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    string cid;
                    if (i != this.dGV_Order.Rows.Count - 2)
                    {
                        for (int k = i; k >= 1; k--)
                        {
                            //被选中状态则删除
                            if (this.dGV_Order.Rows[k - 1].Selected == true)
                            {
                                cid = this.dGV_Order.Rows[k - 1].Cells[0].Value.ToString();
                                dGV_Order.Rows.Remove(dGV_Order.Rows[k - 1]);

                                delOrder = "delete from Component_Order where cy_cID = '" + cid + "'";
                                del.ExecuteUpdate(delOrder);
                            }
                        }
                    }
                    //逆序删除防止串行
                    MessageBox.Show("删除成功");
                }
                else
                    MessageBox.Show("取消删除所选零件");
            }
            catch (Exception)
            {
                MessageBox.Show("未成功删除");
            }
        }

        /// <summary>
        /// 查指定订货日期范围 导出历史数据
        /// </summary>
        private void btnSearchRecord_Click(object sender, EventArgs e)
        {
            string startTime = dTP_oStartTime.Value.ToString("yyyy-MM-dd");
            string endTime = dTP_oEndTime.Value.ToString("yyyy-MM-dd") + "24:00:00";
            string condition;
            ConnSQL con = new ConnSQL();
            DataTable ioInfo = new DataTable();

            condition = "SELECT cy_cID,cy_sID,cy_orderNum,cy_orderTime FROM Component_Order WHERE cy_orderTime BETWEEN '" + startTime + "' AND '" + endTime + "' ORDER BY cy_orderTime DESC";
            ioInfo = con.ExecuteQuery(condition);
            dGV_reportAll.DataSource = ioInfo;

            dGV_reportAll.Columns["cy_cID"].HeaderText = "零件编号";
            dGV_reportAll.Columns["cy_sID"].HeaderText = "供应商编号";
            dGV_reportAll.Columns["cy_orderNum"].HeaderText = "订货数量";
            dGV_reportAll.Columns["cy_orderTime"].HeaderText = "订货时间";

            if (dGV_reportAll.Rows.Count == 0)
                MessageBox.Show("查询时间段内无订货记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }
        private void btnOrderReport2_Click(object sender, EventArgs e)
        {

            if (dGV_reportAll.Rows.Count == 0)
            {
                MessageBox.Show("当前无数据导出！请先刷新今日订货信息，确认当前有无当日订货信息。");
            }
            else
            {
                string Date = DateTime.Now.ToString("yyyy-MM-dd");
                SaveFileDialog s1 = new SaveFileDialog();

                s1.Title = "请选择要导出的位置";
                s1.Filter = "Excel文件|*.xls";
                s1.FileName = dTP_oStartTime.Value.ToShortDateString().Replace('/','-') +"至"+ dTP_oEndTime.Value.ToShortDateString().Replace('/', '-') + "历史订货信息";
                if (s1.ShowDialog() == DialogResult.OK)
                {
                    NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                    NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet(Date);
                    NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(0);

                    row1.CreateCell(0).SetCellValue("订货序号");
                    row1.CreateCell(1).SetCellValue("零件编号");
                    row1.CreateCell(2).SetCellValue("供应商编号");
                    row1.CreateCell(3).SetCellValue("订货数量");
                    row1.CreateCell(4).SetCellValue("订货时间");

                    for (int i = 0; i < dGV_reportAll.Rows.Count; i++)
                    {
                        NPOI.SS.UserModel.IRow rowTemp = sheet1.CreateRow(i + 1);
                        rowTemp.CreateCell(0).SetCellValue(i + 1);
                        rowTemp.CreateCell(1).SetCellValue(dGV_reportAll.Rows[i].Cells[0].Value.ToString());
                        rowTemp.CreateCell(2).SetCellValue(dGV_reportAll.Rows[i].Cells[1].Value.ToString());
                        rowTemp.CreateCell(3).SetCellValue(dGV_reportAll.Rows[i].Cells[2].Value.ToString());
                        rowTemp.CreateCell(4).SetCellValue(dGV_reportAll.Rows[i].Cells[3].Value.ToString());
                    }

                    FileStream orderRepo = File.OpenWrite(s1.FileName.ToString());
                    try
                    {
                        book.Write(orderRepo);
                        orderRepo.Seek(0, SeekOrigin.Begin);
                        MessageBox.Show("导出成功！", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("导出失败！", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        throw;
                    }
                    finally
                    {
                        if (orderRepo != null)
                        {
                            orderRepo.Close();
                        }
                    }
                }
            }
        }

        private void mainForm_Load(object sender, EventArgs e)
        {
            /**
            *
            * 1. 连数据库
            * 2. 查询股票表
            * 3. 查询结果返回给DataGridView控件
            */
        }

        private void btn_baobiao_Click(object sender, EventArgs e)
        {
            if (dGV_reportAll.Rows.Count < 1)
            {
                MessageBox.Show("无报表信息，请先查询。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {

            }
        }

        private void btn_designBangdan_Click(object sender, EventArgs e)
        {
            Report dreport = new Report();
            dreport.Load("F:\\1_University\\Project_C#\\reportTemplate\\bangdan.frx");
            dreport.Design();
            //dreport.Show();

        }
    }
}
