using System;
using System.Collections.Generic;

using System.Text;

using System.Data.SqlClient;
using System.Data;

namespace Kd_Test
{
    class KdSql
    {
        public static void Save_Sql(string sql)
        {
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
            SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
        }

        public static void Cancel(string str,string num)
        {
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
            conn.Open();
            SqlCommand cmd = new SqlCommand("cancle_KD", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(str, SqlDbType.VarChar,50);
            cmd.Parameters[str].Direction = ParameterDirection.Input;
            cmd.Parameters[str].Value = num;
            cmd.ExecuteNonQuery();
            conn.Close();
        }
    
        public static void Kd_Log(string str, string postuser,string str1,string errorcode,string str2,string machinename,string str3,string jsonal)
        {
            SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=sa;pwd=flybarrier");
            conn.Open();
            SqlCommand cmd = new SqlCommand("KD_LOGData", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.Add(str,SqlDbType.VarChar);
            cmd.Parameters[str].Direction = ParameterDirection.Input;
            cmd.Parameters[str].Value = postuser;

            cmd.Parameters.Add(str1, SqlDbType.VarChar);
            cmd.Parameters[str1].Direction = ParameterDirection.Input;
            cmd.Parameters[str1].Value = errorcode;

            cmd.Parameters.Add(str2, SqlDbType.VarChar);
            cmd.Parameters[str2].Direction = ParameterDirection.Input;
            cmd.Parameters[str2].Value = machinename;

            cmd.Parameters.Add(str3, SqlDbType.VarChar);
            cmd.Parameters[str3].Direction = ParameterDirection.Input;
            cmd.Parameters[str3].Value = jsonal;
          
            cmd.ExecuteNonQuery();
            conn.Close();
        }

    }
}
