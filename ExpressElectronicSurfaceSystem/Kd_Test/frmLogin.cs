using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;

using System.Windows.Forms;
using System.Data.SqlClient;
using System.Security.Cryptography;


namespace Kd_Test
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();
            
        }

        private void btn_login_Click(object sender, EventArgs e)
        {
            String name = this.txt_user.Text; // 获取里面的值     
            String password = this.txt_pwd.Text;

            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
            string sqluser = string.Format(@"select * from KD_user where loginid ='{0}'",txt_user.Text);
            SqlDataAdapter sda = new SqlDataAdapter(sqluser, conn);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt != null && dt.Rows.Count > 0)
            {
                loginfo.userName = dt.Rows[0]["lastname"].ToString();
                loginfo.userDepartMent = dt.Rows[0]["departmentname"].ToString();
                loginfo.userID = dt.Rows[0]["id"].ToString();
                if(MD5(txt_pwd.Text)==dt.Rows[0]["password"].ToString())
                    this.DialogResult = DialogResult.OK;
                else
                {
                    MessageBox.Show("您输入的用户名或密码错误！", "登录失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txt_user.Text = "";
                    txt_pwd.Text = "";
                }
            }
            
            else
            { 
                MessageBox.Show("您输入的用户名或密码错误！", "登录失败！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt_user.Text = "";
                txt_pwd.Text = "";
            }          
        }

        private void btn_logout_Click(object sender, EventArgs e)
        {
            Application.Exit();
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
    }
}
