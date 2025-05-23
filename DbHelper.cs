using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace SendOperationPlan
{
    public class DbHelper
    {
        // 数据库连接字符串（建议从配置文件读取）
        //private static string _connectionString = "Server=155.155.100.107;Database=docare;User Id=zj;Password=zj12345678;";
        //private static string _connectionString = "Server=155.155.100.88;Database=CISDB;User Id=zhangjian;Password=Zj@admin;";

        // 从配置文件中读取连接字符串
        private static string _connectionString =
            ConfigurationManager.ConnectionStrings["DB"].ConnectionString;

        //改成配置

        /// <summary>
        /// 执行 SQL 查询或存储过程，返回 DataTable
        /// </summary>
        /// <param name="sqlQuery">SQL 语句或存储过程名</param>
        /// <param name="commandType">命令类型（Text/ StoredProcedure）</param>
        /// <param name="parameters">参数集合</param>
        /// <returns>包含查询结果的 DataTable</returns>
        public static DataTable GetData(string sqlQuery, CommandType commandType = CommandType.Text, SqlParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        command.CommandType = commandType;

                        // 添加参数
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                // 记录日志或抛出异常
                Console.WriteLine($"数据库错误: {ex.Message}");
                throw;
            }

            return dataTable;
        }




        public static int ExecuteNonQuery(string sql,  params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    return cmd.ExecuteNonQuery(); // 返回受影响行数
                }
            }
        }


    }
}
