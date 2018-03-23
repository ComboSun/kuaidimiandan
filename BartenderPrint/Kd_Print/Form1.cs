using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;

using System.Windows.Forms;
using System.Data.SqlClient;
using LitJson;
using Seagull.BarTender.Print;
using System.Drawing.Printing;

namespace Kd_Print
{
    public partial class Form1 : Form
    {
        SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
        BackResult br = new BackResult();
        kdinfo ki = new kdinfo();
        public Form1()
        {
            InitializeComponent();
        }

        public void Info_Select()
        {
            string mon;
            string dy;
            string year = dateTimePicker1.Value.Year.ToString();
            int month = dateTimePicker1.Value.Month;
            if (month < 10) { mon = "0" + month.ToString(); } else { mon = month.ToString(); }
            int day = dateTimePicker1.Value.Day;
            if (day < 10) { dy = "0" + day.ToString(); } else { dy = day.ToString(); }
            string date = year + "-" + mon + "-" + dy;
    
            string sql = string.Format(@"select b.kuaidinum '快递单号',a.send_name '发件人姓名',a.send_mobil '发件人手机',a.send_tel '发件人电话',a.send_printaddr '发件人地址',a.send_company '发件人公司',a.rec_name '收件人姓名',a.rec_mobil '收件人手机',a.rec_tel '收件人电话',a.rec_printaddr '收件人地址',a.rec_company '收件人公司',a.createuser '创建人',b.jsonval,a.printdate from KD_detail a
inner join KD_detailentry b on a.ID=b.Finterid where convert(varchar(10),a.printdate,120)='{0}'and a.createuser ='{1}' and b.kuaidinum<>'' order by a.printdate desc", date, txt_input.Text);
            SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.ColumnHeadersHeight =30;
               // this.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("微软雅黑", 10);
                //设置标题内容居中显示;  `
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

                //设置列宽
                dataGridView1.Columns[0].Width = 80;
                dataGridView1.Columns[1].Width = 80;
                dataGridView1.Columns[2].Width = 80;
                dataGridView1.Columns[3].Width = 80;
                dataGridView1.Columns[4].Width = 140;
                dataGridView1.Columns[6].Width = 80;
                dataGridView1.Columns[7].Width = 80;
                dataGridView1.Columns[8].Width = 80;
                dataGridView1.Columns[9].Width = 140;
                dataGridView1.Columns[10].Width = 80;
                dataGridView1.Columns[11].Width = 80;
                
                this.dataGridView1.RowHeadersVisible = false;//datagridview前面的空白部分去除
                dataGridView1.Columns[12].Visible = false;
                dataGridView1.Columns[13].Visible = false;
            }            
        }

        public void Parase_Json()
        {
            JsonData jd = JsonMapper.ToObject(ki.jsonal);
            //send
            JsonData jdsend = jd["send"];
            ki.kuaidicom = (string)jdsend["kuaidicom"];
            JsonData jdrecman = jdsend["recMan"];
            ki.Rec_name = (string)jdrecman["name"];
            ki.Rec_mobile = (string)jdrecman["mobile"];
            ki.Rec_tel = (string)jdrecman["tel"];
            ki.Rec_paddr = (string)jdrecman["printAddr"];
            JsonData jdsendman = jdsend["sendMan"];
            ki.Send_name = (string)jdsendman["name"];
            ki.Send_mobile = (string)jdsendman["mobile"];
            ki.Send_tel = (string)jdsendman["tel"];
            ki.Send_paddr = (string)jdsendman["printAddr"];
            ki.weight = (string)jdsend["weight"];
            ki.Send_remark=(string)jdsend["remark"];
            //receive
            JsonData jsrecive = jd["recive"];
            br.result = (bool)jsrecive["result"];
            br.message = (string)jsrecive["message"];
            br.status = (string)jsrecive["status"];
            JsonData jdItems = jsrecive["data"];
            for (int i = 0; i < jdItems.Count; i++)
            {
                br.kuaidinum = (string)jdItems[i]["kuaidinum"];
                if (ki.jsonal.Contains("returnNum"))
                    br.returnNum = (string)jdItems[i]["returnNum"];
                else
                    br.returnNum = "";
                if (ki.jsonal.Contains("childNum"))
                    br.childNum = (string)jdItems[i]["childNum"];
                else
                    br.childNum = "";
                if (ki.jsonal.Contains("bulkpen"))
                    br.bulkpen = (string)jdItems[i]["bulkpen"];
                else
                    br.bulkpen = "";
                if (ki.jsonal.Contains("orgCode"))
                    br.orgCode = (string)jdItems[i]["orgCode"];
                else
                    br.orgCode = "";
                if (ki.jsonal.Contains("orgName"))
                    br.orgName = (string)jdItems[i]["orgName"];
                else
                    br.orgName = "";
                if (ki.jsonal.Contains("destCode"))
                    br.destCode = (string)jdItems[i]["destCode"];
                else
                    br.destCode = "";
                if (ki.jsonal.Contains("destName"))
                    br.destName = (string)jdItems[i]["destName"];
                else
                    br.destName = "";
                if (ki.jsonal.Contains("orgSortingCode"))
                    br.orgSortingCode = (string)jdItems[i]["orgSortingCode"];
                else
                    br.orgSortingCode = "";
                if (ki.jsonal.Contains("orgSortingName"))
                    br.orgSortingName = (string)jdItems[i]["orgSortingName"];
                else
                    br.orgSortingName = "";
                if (ki.jsonal.Contains("destSortingCode"))
                    br.destSortingCode = (string)jdItems[i]["destSortingCode"];
                else
                    br.destSortingCode = "";
                if (ki.jsonal.Contains("destSortingName"))
                    br.destSortingName = (string)jdItems[i]["destSortingName"];
                else
                    br.destSortingName = "";
                if (ki.jsonal.Contains("orgExtra"))
                    br.orgExtra = (string)jdItems[i]["orgExtra"];
                else
                    br.orgExtra = "";
                if (ki.jsonal.Contains("destExtra"))
                    br.destExtra = (string)jdItems[i]["destExtra"];
                else
                    br.destExtra = "";
                if (ki.jsonal.Contains("pkgCode"))
                    br.pkgCode = (string)jdItems[i]["pkgCode"];
                else
                    br.pkgCode = "";
                if (ki.jsonal.Contains("pkgName"))
                    br.pkgName = (string)jdItems[i]["pkgName"];
                else
                    br.pkgName = "";
                if (ki.jsonal.Contains("road"))
                    br.road = (string)jdItems[i]["road"];
                else
                    br.road = "";
                if (ki.jsonal.Contains("qrCode"))
                    br.qrCode = (string)jdItems[i]["qrCode"];
                else
                    br.qrCode = "";
                if (ki.jsonal.Contains("orderNum"))
                    br.orderNum = (string)jdItems[i]["orderNum"];
                else
                    br.orderNum = "";
                if (ki.jsonal.Contains("expressCode"))
                    br.expressCode = (string)jdItems[i]["expressCode"];
                else
                    br.expressCode = "";
                if (ki.jsonal.Contains("expressName"))
                    br.expressName = (string)jdItems[i]["expressName"];
                else
                    br.expressName = "";
                br.templateurl = (string)jdItems[i]["templateurl"][0];
                br.template = (string)jdItems[i]["template"][0];
            }          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            //int a = dataGridView1.CurrentRow.Index;
            //ki.jsonal = dataGridView1.Rows[a].Cells[12].Value.ToString();
            //textBox1.Text = ki.jsonal;
        }

        private void btn_select_Click(object sender, EventArgs e)
        {
            Info_Select();
        }

        //打印标签
        public void Print_Labels()
        {
            DateTime dt = DateTime.Now;
            string print_time = dt.ToString();
            Engine engine = new Engine();
            engine.Start();
            LabelFormatDocument format = null;
            ki.OldActPrn = cmbPrinterList.Text;

            //采用中通发货
            if (ki.kuaidicom == "zhongtong")
            {
                try
                {
                    format = engine.Documents.Open(Environment.CurrentDirectory + "\\zt.btw");//@"C:\Users\sunkb.CHENZHU\Desktop\zt.btw"
                    format.SubStrings["bulkpen"].Value = br.bulkpen;
                    //format.SubStrings["pkgname"].Value = br.pkgName;
                    //format.SubStrings["destName"].Value = br.destName;
                    //format.SubStrings["orgSortingCode"].Value = br.orgSortingCode;
                    format.SubStrings["orgname"].Value = br.orgName;
                    format.SubStrings["recname"].Value = ki.Rec_name;
                    if (ki.Rec_mobile == "")
                        format.SubStrings["recphone"].Value = ki.Rec_tel;
                    else
                        format.SubStrings["recphone"].Value = ki.Rec_mobile;
                    format.SubStrings["recaddress"].Value = ki.Rec_paddr;
                    format.SubStrings["sendname"].Value = ki.Send_name;
                    format.SubStrings["sendphone"].Value = ki.Send_tel;
                    format.SubStrings["sendaddress"].Value = ki.Send_paddr;
                    format.SubStrings["barcode"].Value = br.kuaidinum;
                    format.SubStrings["remark"].Value = ki.Send_remark;
                    if (ki.Rec_mobile != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_mobile.Substring(ki.Rec_mobile.Length - 4, 4);
                    }
                    else if (ki.Rec_tel != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_tel.Substring(ki.Rec_tel.Length - 4, 4);
                    }
                    if (br.qrCode != "")
                        format.SubStrings["qrcode"].Value = br.qrCode;
                    //format.PrintSetup.NumberOfSerializedLabels = 1;   //设置打印份数
                    //format.PrintSetup.IdenticalCopiesOfLabel = 1;
                    format.PrintSetup.PrinterName = ki.OldActPrn;
                    format.Print();
                    engine.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message, "Error");
                }
            }

            //采用申通发货
            if (ki.kuaidicom == "shentong")
            {
                try
                {
                    format = engine.Documents.Open(Environment.CurrentDirectory + "\\st.btw");//@"C:\Users\sunkb.CHENZHU\Desktop\st.btw"
                    format.SubStrings["bulkpen"].Value = br.bulkpen;
                    format.SubStrings["pkgname"].Value = br.pkgName;
                    //format.SubStrings["destName"].Value = br.destName;
                    //format.SubStrings["orgSortingCode"].Value = br.orgSortingCode;
                    format.SubStrings["recname"].Value = ki.Rec_name;
                    if (ki.Rec_mobile == "")
                        format.SubStrings["recphone"].Value = ki.Rec_tel;
                    else
                        format.SubStrings["recphone"].Value = ki.Rec_mobile;
                    format.SubStrings["recaddress"].Value = ki.Rec_paddr;
                    format.SubStrings["sendname"].Value = ki.Send_name;
                    format.SubStrings["sendphone"].Value = ki.Send_tel;
                    format.SubStrings["sendaddress"].Value = ki.Send_paddr;
                    format.SubStrings["barcode"].Value = br.kuaidinum;
                    format.SubStrings["remark"].Value = ki.Send_remark;
                    //if (ki.Rec_mobile != "")
                    //{
                    //    format.SubStrings["phonecut"].Value = ki.Rec_mobile.Substring(ki.Rec_mobile.Length - 4, 4);
                    //}
                    //else if (ki.Rec_tel != "")
                    //{
                    //    format.SubStrings["phonecut"].Value = ki.Rec_tel.Substring(ki.Rec_tel.Length - 4, 4);
                    //}
                    if (br.qrCode != "")
                        format.SubStrings["qrcode"].Value = br.qrCode;
                    //format.PrintSetup.NumberOfSerializedLabels = 1;   //设置打印份数
                    //format.PrintSetup.IdenticalCopiesOfLabel = 1;
                    format.PrintSetup.PrinterName = ki.OldActPrn;
                    format.Print();
                    engine.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message, "Error");
                }
            }

            //采用顺丰发货   
            if (ki.kuaidicom == "shunfeng")
            {
                try
                {
                    format = engine.Documents.Open(Environment.CurrentDirectory + "\\sf.btw");
                    format.SubStrings["destcode"].Value = br.destCode;
                    if (ki.Rec_mobile == "")
                        format.SubStrings["rec"].Value = ki.Rec_name + " " + ki.Rec_tel + " " + ki.Rec_paddr;
                    else
                        format.SubStrings["rec"].Value = ki.Rec_name + " " + ki.Rec_mobile + " " + ki.Rec_paddr;
                    format.SubStrings["send"].Value = ki.Send_name + " " + ki.Send_tel + " " + ki.Send_paddr;
                    format.SubStrings["barcode"].Value = br.kuaidinum;
                    format.SubStrings["remark"].Value = ki.Send_remark;
                    format.SubStrings["outdate"].Value = print_time;
                    if (ki.Rec_mobile != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_mobile.Substring(ki.Rec_mobile.Length - 4, 4);
                    }
                    else if (ki.Rec_tel != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_tel.Substring(ki.Rec_tel.Length - 4, 4);
                    }
                    format.PrintSetup.PrinterName = ki.OldActPrn;
                    format.Print();
                    engine.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message, "Error");
                }
            }

            //采用韵达发货
            if (ki.kuaidicom == "yunda")
            {
                try
                {
                    format = engine.Documents.Open(Environment.CurrentDirectory + "\\yd.btw");
                    format.SubStrings["bulkpen"].Value = br.bulkpen;
                    format.SubStrings["pkgname"].Value = br.pkgName;
                    format.SubStrings["destName"].Value = br.destName;
                    format.SubStrings["orgSortingCode"].Value = br.orgSortingCode;
                    format.SubStrings["recname"].Value = ki.Rec_name;
                    if (ki.Rec_mobile == "")
                        format.SubStrings["recphone"].Value = ki.Rec_tel;
                    else
                        format.SubStrings["recphone"].Value = ki.Rec_mobile;
                    format.SubStrings["recaddress"].Value = ki.Rec_paddr;
                    format.SubStrings["sendname"].Value = ki.Send_name;
                    format.SubStrings["sendphone"].Value = ki.Send_tel;
                    format.SubStrings["sendaddress"].Value = ki.Send_paddr;
                    format.SubStrings["barcode"].Value = br.kuaidinum;
                    format.SubStrings["remark"].Value = ki.Send_remark;
                    if (ki.Rec_mobile != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_mobile.Substring(ki.Rec_mobile.Length - 4, 4);
                    }
                    else if (ki.Rec_tel != "")
                    {
                        format.SubStrings["phonecut"].Value = ki.Rec_tel.Substring(ki.Rec_tel.Length - 4, 4);
                    }
                    if (br.qrCode != "")
                        format.SubStrings["qrcode"].Value = br.qrCode;
                    //format.PrintSetup.NumberOfSerializedLabels = 1;   //设置打印份数
                    //format.PrintSetup.IdenticalCopiesOfLabel = 1;
                    format.PrintSetup.PrinterName = ki.OldActPrn;
                    format.Print();
                    engine.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message, "Error");
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ReadOnly = true;
            Get_PrinterName();
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex >= 0)
                {
                    //若行已是选中状态就不再进行设置
                    if (dataGridView1.Rows[e.RowIndex].Selected == false)
                    {
                        dataGridView1.ClearSelection();
                        dataGridView1.Rows[e.RowIndex].Selected = true;
                    }
                    //只选中一行时设置活动单元格
                    if (dataGridView1.SelectedRows.Count == 1)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    }
                    //弹出操作菜单
                    contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                }
            }
        }

        private void 打印ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //int a = dataGridView1.CurrentRow.Index;
            //string kuaidinum = dataGridView1.Rows[a].Cells[0].Value.ToString();    
       
            //ki.jsonal = dataGridView1.Rows[a].Cells[12].Value.ToString();

            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                ki.jsonal = row.Cells[12].Value.ToString();
                //MessageBox.Show(ki.jsonal);
                string kuaidinum = row.Cells[0].Value.ToString();
                if (kuaidinum != null)
                {
                    Parase_Json();

                    Print_Labels();
                }                  
            }

            //foreach (DataGridViewRow gdvr in dataGridView1.SelectedRows)
            //{
            //    ki.jsonal= gdvr.Cells[12].Value.ToString();
            //    Parase_Json();
             
            //    Print_Labels();
            //}


            //if (kuaidinum != "")
            //{
            //    Parase_Json();
            //    Print_Labels();
            //}
            //else
            //{
            //    MessageBox.Show("无此快递单号！请重新选择！", "打印失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //}
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.Font = new Font("微软雅黑", 10);
        }

        public void Get_PrinterName()
        {
            int i;
            string pkInstalledPrinters;
            using (PrintDocument pd = new PrintDocument())
            {
                for (i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)  //开始遍历  
                {
                    pkInstalledPrinters = PrinterSettings.InstalledPrinters[i];  //取得名称  
                    cmbPrinterList.Items.Add(pkInstalledPrinters);               //加入ComboBox  
                    //if (pd.PrinterSettings.IsDefaultPrinter)                     //判断是否为默认打印机  
                    //{
                    //    ki.OldActPrn = pd.PrinterSettings.PrinterName;              //保存名称，后面要用  
                    //    cmbPrinterList.Text = pd.PrinterSettings.PrinterName;    //显示默认打印机名称  
                    //}
                }
                cmbPrinterList.Text = pd.PrinterSettings.PrinterName;                
            } 
        } 
    }
}
