using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data;

namespace MyExcel.Models
{
    class DBViewConnect
    {
        public DBViewConnect() { }
        private static string connectionString = @"Data Source=(LocalDB)\mssqllocaldb;Initial Catalog=ProgData;Integrated Security=True;Pooling=False";

        public static void UpdateToViewDB(string row, string col, string value)
        {
            //string sqlExpression = $"UPDATE Table SET {col}='{value}' WHERE id={row};";
            //using (SqlConnection connection = new SqlConnection(connectionString))
            //{
            //    connection.Open();
            //    // Создаем объект DataAdapter
            //    SqlDataAdapter adapter = new SqlDataAdapter(sqlExpression, connection);
            //    // Создаем объект Dataset
            //    DataSet ds = new DataSet();
            //    // Заполняем Dataset
            //    adapter.Fill(ds);
            //}
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string sqlExpression = $"UPDATE ViewTable SET content='{value}' WHERE position='{row + " " + col}';";
                

                SqlCommand cmd = new SqlCommand(sqlExpression, connection);
                try
                {
                    connection.Open();
                    if(cmd.ExecuteNonQuery() <= 0)
                    {
                        sqlExpression = $"INSERT INTO ViewTable (position, content) VALUES ('{row + " " + col}', '{value}');";
                        cmd = new SqlCommand(sqlExpression, connection);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (SqlException e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
    }
}
