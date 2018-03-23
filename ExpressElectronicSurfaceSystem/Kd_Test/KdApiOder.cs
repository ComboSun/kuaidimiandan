using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using LitJson;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Data.SqlClient;

namespace Kd_Test
{
    public partial class KDApiOder : Form
    {
        //快递100-电子面单接口（API）的授权密匙(Key)
        private string key = "TuZxCiov5623";
        //快递100-电子面单接口（API）的secret
        private string secret = "ff08e600-709a-4692-89ca-dc890a0e2b00";
        //请求URL
        private string url = "http://api.kuaidi100.com/eorderapi.do?method=getElecOrder";
        //连接数据库
        //SqlConnection conn = new SqlConnection("server=192.168.1.6;database=Exd2TEST;uid=sa;pwd=flybarrier");

        BackResult br = new BackResult();
        string param = ""; //param 发送快递信息的json转字符串值
        string guid = "";  //guid作为数据表的id号
        string result = "";//声明变量，返回结果result  
        //string kdcompany = "";

        public KDApiOder()
        {
            InitializeComponent();
        }
        
        /// <summary>
        /// Json方式  电子面单
        /// </summary>
        /// <returns></returns>
        //发送方式：json
        public void Kd_Send()
        {
            KdService ks = new KdService();
            //如果选择中通快递
            if (rb_zt.Checked) { KdService.kdcompany = "zhongtong"; KdService.kdpartnerid = "02109.456"; KdService.kdpartnerkey = "3135456378"; KdService.kdnet = "02109"; }
            //如果选择顺丰快递
            if (rb_sf.Checked) { KdService.kdcompany = "shunfeng"; KdService.kdpartnerid = "0210401002"; KdService.kdpartnerkey = ""; KdService.kdnet = ""; }
            //如果选择韵达快递
            //if (rb_yd.Checked) { KdService.kdcompany = "yunda"; KdService.kdpartnerid = "2015456006"; KdService.kdpartnerkey = "rguSvhT6bC23KiNmHRAd8UVeqZpYyQ"; KdService.kdnet = "201545"; }
            //如果选择申通快递
            if (rb_st.Checked) { KdService.kdcompany = "shentong"; KdService.kdpartnerid = "辰竹仪表"; KdService.kdpartnerkey = "Sto*00159"; KdService.kdnet = "上海松江公司"; }

            StringBuilder sb = new StringBuilder();
            //组合param参数
            JsonWriter writer = new JsonWriter(sb);
            writer.WriteObjectStart();
            writer.WritePropertyName("partnerId");
            writer.Write(KdService.kdpartnerid);
            writer.WritePropertyName("partnerKey");
            writer.Write(KdService.kdpartnerkey);
            writer.WritePropertyName("net");
            writer.Write(KdService.kdnet);
            writer.WritePropertyName("kuaidicom");
            writer.Write(KdService.kdcompany);
            writer.WritePropertyName("kuaidinum");
            writer.Write("");
            writer.WritePropertyName("orderId");
            writer.Write("");

            //定义收件人
            writer.WritePropertyName("recMan");
            writer.WriteObjectStart();
            writer.WritePropertyName("name");
            writer.Write(Rec_name);
            writer.WritePropertyName("mobile");
            writer.Write(Rec_mobile.Trim());
            writer.WritePropertyName("tel");
            writer.Write(Rec_tel.Trim());
            writer.WritePropertyName("zipCode");
            writer.Write("");
            writer.WritePropertyName("province");
            writer.Write("");
            writer.WritePropertyName("city");
            writer.Write("");
            writer.WritePropertyName("district");
            writer.Write("");
            writer.WritePropertyName("addr");
            writer.Write("");
            writer.WritePropertyName("printAddr");
            writer.Write(Rec_paddr);
            writer.WritePropertyName("company");
            writer.Write(Rec_company);
            writer.WriteObjectEnd();

            //定义发件人
            writer.WritePropertyName("sendMan");
            writer.WriteObjectStart();
            writer.WritePropertyName("name");
            writer.Write(Send_name);
            writer.WritePropertyName("mobile");
            writer.Write(Send_mobile.Trim());
            writer.WritePropertyName("tel");
            writer.Write(Send_tel.Trim());
            writer.WritePropertyName("zipCode");
            writer.Write("");//邮编
            writer.WritePropertyName("province");
            writer.Write("");
            writer.WritePropertyName("city");
            writer.Write("");
            writer.WritePropertyName("district");
            writer.Write("");
            writer.WritePropertyName("addr");
            writer.Write("");
            writer.WritePropertyName("printAddr");
            writer.Write(Send_paddr);
            writer.WritePropertyName("company");
            writer.Write(Send_company);
            writer.WriteObjectEnd();

            //物品信息及其他参数
            writer.WritePropertyName("cargo");
            writer.Write("");
            writer.WritePropertyName("count");
            writer.Write("1");
            writer.WritePropertyName("weight");
            writer.Write(Send_weight);
            writer.WritePropertyName("volumn");
            writer.Write("");
            writer.WritePropertyName("payType");//寄件方法
            writer.Write("SHIPPER");
            writer.WritePropertyName("expType");
            writer.Write("标准快递");
            writer.WritePropertyName("remark");//备注
            writer.Write(Send_remark);
            writer.WritePropertyName("valinsPay");
            writer.Write("0");
            writer.WritePropertyName("collection");
            writer.Write("0");
            writer.WritePropertyName("needChild");
            writer.Write("0");
            writer.WritePropertyName("needBack");
            writer.Write("0");
            writer.WritePropertyName("needTemplate");//返回模板
            writer.Write("1");
            writer.WriteObjectEnd();
            param = sb.ToString();

            //获取13位时间戳 t
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            string t = Convert.ToInt64(ts.TotalMilliseconds).ToString();

            //获取加电商签名sign，通过MD5加密
            string sendJson = param + t + key + secret;
            string sign = MD5(sendJson);

            IDictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("param", param);
            parameters.Add("sign", sign);
            parameters.Add("t", t);
            parameters.Add("key", key);

            //Post方式提交数据，发送HTTP请求
            Encoding encoding = Encoding.GetEncoding("utf-8");
            HttpWebResponse response = CreatePostHttpResponse(url, parameters, encoding);
            Stream stream = response.GetResponseStream();   //获取响应的字符串流  
            StreamReader sr = new StreamReader(stream); //创建一个stream读取流  
            result = sr.ReadToEnd();
            //try
            //{
            //    FileStream fs = new FileStream(@"C:\Users\sunkb.CHENZHU\Desktop\Receive_JSON.txt", FileMode.OpenOrCreate);
            //    StreamWriter sw = new StreamWriter(fs);
            //    sw.WriteLine(result);
            //    sw.Close();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error:" + ex.Message, "Error");
            //}
        }

        public void Prase_JSON()
        {
            //1.发送json并获取返回的json
            Kd_Send();
            //try
            //{
            //    FileStream fs = new FileStream(@"C:\Users\sunkb.CHENZHU\Desktop\Receive_JSON.txt", FileMode.Open);//初始化文件流
            //    byte[] array = new byte[fs.Length];//初始化字节数组
            //    fs.Read(array, 0, array.Length);//读取流中数据到字节数组中
            //    fs.Close();//关闭流
            //    result = Encoding.UTF8.GetString(array);//将字节数组转化为字符串                
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error:" + ex.Message, "Error");
            //}
            string jsonal = @"{""send"":" + param + ",";
            jsonal += @"""recive"":" + result + "}";           

            //2.解析返回结果的JSON     
            JsonData jd = JsonMapper.ToObject(result);
            br.result = (bool)jd["result"];
            br.message = (string)jd["message"];
            br.status = (string)jd["status"];
            JsonData jdItems = jd["data"];
            if (br.status == "200" && br.result == true)
            {
                for (int i = 0; i < jdItems.Count; i++)
                {
                    br.kuaidinum = (string)jdItems[i]["kuaidinum"];
                    if (result.Contains("returnNum"))
                        br.returnNum = (string)jdItems[i]["returnNum"];
                    else
                        br.returnNum = "";
                    if (result.Contains("childNum"))
                        br.childNum = (string)jdItems[i]["childNum"];
                    else
                        br.childNum = "";
                    if (result.Contains("bulkpen"))
                        br.bulkpen = (string)jdItems[i]["bulkpen"];
                    else
                        br.bulkpen = "";
                    if (result.Contains("orgCode"))
                        br.orgCode = (string)jdItems[i]["orgCode"];
                    else
                        br.orgCode = "";
                    if (result.Contains("orgName"))
                        br.orgName = (string)jdItems[i]["orgName"];
                    else
                        br.orgName = "";
                    if (result.Contains("destCode"))
                        br.destCode = (string)jdItems[i]["destCode"];
                    else
                        br.destCode = "";
                    if (result.Contains("destName"))
                        br.destName = (string)jdItems[i]["destName"];
                    else
                        br.destName = "";
                    if (result.Contains("orgSortingCode"))
                        br.orgSortingCode = (string)jdItems[i]["orgSortingCode"];
                    else
                        br.orgSortingCode = "";
                    if (result.Contains("orgSortingName"))
                        br.orgSortingName = (string)jdItems[i]["orgSortingName"];
                    else
                        br.orgSortingName = "";
                    if (result.Contains("destSortingCode"))
                        br.destSortingCode = (string)jdItems[i]["destSortingCode"];
                    else
                        br.destSortingCode = "";
                    if (result.Contains("destSortingName"))
                        br.destSortingName = (string)jdItems[i]["destSortingName"];
                    else
                        br.destSortingName = "";
                    if (result.Contains("orgExtra"))
                        br.orgExtra = (string)jdItems[i]["orgExtra"];
                    else
                        br.orgExtra = "";
                    if (result.Contains("destExtra"))
                        br.destExtra = (string)jdItems[i]["destExtra"];
                    else
                        br.destExtra = "";
                    if (result.Contains("pkgCode"))
                        br.pkgCode = (string)jdItems[i]["pkgCode"];
                    else
                        br.pkgCode = "";
                    if (result.Contains("pkgName"))
                        br.pkgName = (string)jdItems[i]["pkgName"];
                    else
                        br.pkgName = "";
                    if (result.Contains("road"))
                        br.road = (string)jdItems[i]["road"];
                    else
                        br.road = "";
                    if (result.Contains("qrCode"))
                        br.qrCode = (string)jdItems[i]["qrCode"];
                    else
                        br.qrCode = "";
                    if (result.Contains("orderNum"))
                        br.orderNum = (string)jdItems[i]["orderNum"];
                    else
                        br.orderNum = "";
                    if (result.Contains("expressCode"))
                        br.expressCode = (string)jdItems[i]["expressCode"];
                    else
                        br.expressCode = "";
                    if (result.Contains("expressName"))
                        br.expressName = (string)jdItems[i]["expressName"];
                    else
                        br.expressName = "";
                    br.templateurl = (string)jdItems[i]["templateurl"][0];
                    br.template = (string)jdItems[i]["template"][0];
                }
                MessageBox.Show("提交成功！快递单号为：" + br.kuaidinum + ",请前往前台打印！", "发送成功！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Save_log(KdService.kdcompany + "发送正常！单号：" + br.kuaidinum, "", jsonal);
            }
            else
            {
                MessageBox.Show("快递单号获取失败，请重新输入！", "发送失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Save_log("发送数据格式错误！错误代码为："+br.message, "", jsonal);
            }
           
            guid = System.Guid.NewGuid().ToString();

            //写入表KD_detail
            string sqlsend = @"insert into KD_detail (ID,PrintDate,cancle,issend,rec_name,rec_mobil,rec_tel,rec_printaddr,rec_company,send_name,send_mobil,send_tel,send_printaddr,send_company,exptype,remark,needchild,neeBack,needtemplate,createuser,department,kdcompany)
                               values(" + "'" + guid + "','" + DateTime.Now + "','0','0'," + "'" + Rec_name + "','" + Rec_mobile + "','" + Rec_tel + "','" + Rec_paddr + "','" + Rec_company + "','" + Send_name + "','" + Send_mobile + "','" + Send_tel + "','" + Send_paddr + "','" + Send_company + "'," + "'标准快递','" + Send_remark + "'" + ",'0','0','1','" + loginfo.userName + "','" + loginfo.userDepartMent + "','" + KdService.kdcompany + "'" + ")";
            KdSql.Save_Sql(sqlsend);
            //写入数据库，表KD_detailentry
            string sqlre = @"insert into KD_detailentry (Finterid,kuaidinum,bulkpen,orgcode,orgname,destcode,destname,orgsortingcode,orgsortingname,destsortingcode,destsortingname,orgextra,destextra,pkgcode,road,qrcode,orderNum,expresscode,expressname,templater,url,pkgname,jsonval) 
                             values(" + "'" + guid + "','" + br.kuaidinum + "','" + br.bulkpen + "','" + br.orgCode + "','" + br.orgName + "','" + br.destCode + "','" + br.destName + "','" + br.orgSortingCode + "','" + br.orgSortingName + "','" + br.destSortingCode + "','" + br.destSortingName + "','" + br.orgExtra + "','" + br.destExtra + "','" + br.pkgCode + "','" + br.road + "','" + br.qrCode + "','" + br.orderNum + "','" + br.expressCode + "','" + br.expressName + "','" + br.template + "','" + br.templateurl + "','" + br.pkgName + "','" + jsonal+"'" + ")";
            KdSql.Save_Sql(sqlre);
            //Print_Labels();       
        }

        public void Cancel_Kuaidi()
        {
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
            string sqluser = string.Format(@"select * from KD_detailentry where kuaidinum ='{0}'", txt_kuaidinum.Text);
            SqlDataAdapter sda = new SqlDataAdapter(sqluser, conn);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt != null && dt.Rows.Count > 0) 
            {
                KdSql.Cancel("@kuaidinumber", txt_kuaidinum.Text);
                MessageBox.Show("作废成功！", "成功！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Save_log("作废了单据编号为：" + txt_kuaidinum.Text + "的单据！", "", "");
            }                
            else
            {
                MessageBox.Show("作废失败，快递单号不存在！", "作废失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受     
        }

        public static HttpWebResponse CreatePostHttpResponse(string url, IDictionary<string, string> parameters, Encoding charset)
        {
            HttpWebRequest request = null;
            //HTTPSQ请求  
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
            try
            {
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";
                //如果需要POST数据     
                if (!(parameters == null || parameters.Count == 0))
                {
                    StringBuilder buffer = new StringBuilder();
                    int i = 0;
                    foreach (string key in parameters.Keys)
                    {
                        if (i > 0)
                        {
                            buffer.AppendFormat("&{0}={1}", key, parameters[key]);
                        }
                        else
                        {
                            buffer.AppendFormat("{0}={1}", key, parameters[key]);
                        }
                        i++;
                    }
                    byte[] data = charset.GetBytes(buffer.ToString());
                    using (Stream stream = request.GetRequestStream())
                    {
                        stream.Write(data, 0, data.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex.Message, "Error");
            }
            return request.GetResponse() as HttpWebResponse;
        }

        ///<summary>
        /// 字符串MD5加密
        ///</summary>
        ///<param name="str">要加密的字符串</param>
        ///<returns>密文</returns>
        private string MD5(string str)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] InBytes = Encoding.GetEncoding("UTF-8").GetBytes(str);
            byte[] OutBytes = md5.ComputeHash(InBytes);
            string OutString = "";
            for (int i = 0; i < OutBytes.Length; i++)
            {
                OutString += OutBytes[i].ToString("x2");
            }
            return OutString.ToUpper();
        }

        private void btn_kdprint_Click(object sender, EventArgs e)
        {
            if (Rec_name != "" && Rec_paddr != "" && (Rec_mobile != "" || Rec_tel != "") && Send_name != "" && Send_paddr != "" && (Send_mobile != "" || Send_tel != "") && (rb_zt.Checked || rb_st.Checked || rb_sf.Checked))
            {
                Prase_JSON();
            }
            else
            {
                MessageBox.Show("请完善快递信息！", "发送失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }             
        }

        private void KDApiOder_Load(object sender, EventArgs e)
        {
            st_user.Text = loginfo.userName;
            st_userid.Text = loginfo.userID;
            st_dep.Text = loginfo.userDepartMent;
            st_time.Text = DateTime.Now.ToString();      
            timer1.Interval = 1000;
            timer1.Start();
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            st_time.Text = DateTime.Now.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cancel_Kuaidi();       
        }

        public void Save_log(string errorcode, string machinename, string xmlvalues)
        {
            KdSql.Kd_Log("@Postuser", loginfo.userName, "@ErrorCode", errorcode, "@MachineName", machinename, "@XMLValues", xmlvalues);       
        }   
    }
}
