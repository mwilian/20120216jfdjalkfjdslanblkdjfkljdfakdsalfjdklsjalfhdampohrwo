using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using System.Configuration;
using System.Drawing;
using System.IO;
using MySql.Data.MySqlClient;

namespace QueryBuilder
{
    public class CoreConnection
    {
        public static string _connectionString = "";
        protected MySqlConnection connection;
        protected MySqlDataAdapter adapter;
        protected MySqlCommand command;
        protected MySqlTransaction trans;

        public CoreConnection()
        {
            connection = new MySqlConnection(_connectionString + ";Connect Timeout=500");
        }
        public static string ConnectionString
        {
            get
            {
                return _connectionString;
            }
            set
            {
                _connectionString = value;
            }
        }
        public static void SetConnection(string connect)
        {
            _connectionString = connect;
        }
        public bool TestConnect()
        {
            bool flag = true;
            try
            {
                connect();
                disconnect();
            }
            catch
            {
                flag = false;

            }
            return flag;
        }
        public void connect()
        {
            connection = new MySqlConnection(_connectionString);
            command = new MySqlCommand();
            connection.Open();
            command.Connection = connection;
        }

        public bool BeginTransaction(ref String sErr)
        {
            try
            {
                connection = new MySqlConnection(_connectionString);
                command = new MySqlCommand();
                command.Connection = connection;
                trans = connection.BeginTransaction();
                command.Transaction = trans;
                return true;
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
                return false;
            }
        }
        public bool CommitTransaction(ref String sErr)
        {
            try
            {
                trans.Commit();
                return true;
            }
            catch (Exception e)
            {
                try
                {
                    trans.Rollback();
                    sErr = e.Message;
                    return false;
                }
                catch (Exception ex)
                {
                    sErr = ex.Message;
                    return false;
                }
            }
            finally
            {
                connection.Close();
            }
        }
        public bool RollbackTransaction(ref String sErr)
        {

            try
            {
                trans.Rollback();
                return false;
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
                return false;
            }
            finally
            {
                connection.Close();
            }
        }


        /// <summary>
        /// Thêm tham số vào command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        protected void AddParameter(string name, object value)
        {
            MySqlParameter para = command.CreateParameter();
            para.ParameterName = name;
            para.Value = value;
            command.Parameters.Add(para);
        }
        public static Byte[] ConvertImageToByte(Image value)
        {
            // string filename = ".\\Template\\" +Convert.ToString(DateTime.Now.ToFileTime());
            // value.Save(filename);
            // FileInfo fileInfo = new FileInfo(filename);
            // FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite);
            // Byte[] barrImg = new Byte[Convert.ToInt32(fileInfo.Length)];
            // int iBytesRead = fs.Read(barrImg, 0,
            //              Convert.ToInt32(fileInfo.Length));
            //// File.Delete(filename);
            // fs.Close();
            MemoryStream ms = new MemoryStream();
            value.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            Byte[] barrImg = ms.ToArray();
            return barrImg;
        }
        public static Image ConvertByteToImage(Byte[] value)
        {
            //string filename = ".\\Template\\" + Convert.ToString(DateTime.Now.ToFileTime());
            //FileStream fs = new FileStream(filename, FileMode.CreateNew, FileAccess.ReadWrite);
            //fs.Write(value, 0, value.Length);
            //fs.Flush();
            //fs.Close();
            MemoryStream fs = new MemoryStream(value);
            Image kq = Image.FromStream(fs);
            //File.Delete(filename);
            return kq;
        }
        protected void AddParameterImage(string name, System.Drawing.Image value)
        {
            MySqlParameter para = command.CreateParameter();
            para.ParameterName = name;
            Byte[] barrImg = ConvertImageToByte(value);
            para.Value = barrImg;
            command.Parameters.Add(para);
        }

        /// <summary>
        /// Thêm mảng tham số vào command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        protected void AddParameters(MySqlParameter[] arrParam)
        {
            for (int i = 0; i < arrParam.Length; i++)
                command.Parameters.Add(arrParam[i]);
        }

        /// <summary>
        /// Tạo mảng tham số
        /// </summary>
        /// <param name="objEntr"></param>
        /// <returns></returns>
        protected virtual void GetParammeter(object objEntr)
        {
            return;
        }
        protected virtual void InitSPCommand(string strCommandText)
        {
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = strCommandText;
            command.Parameters.Clear();
        }

        public void disconnect()
        {
            command.Dispose();
            connection.Close();
        }

        /// <summary>
        /// Them tham so vao SP_command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>


        public DataTable executeSelectQuery(string sqlString)
        {
            DataSet ds = new DataSet();
            adapter = new MySqlDataAdapter(sqlString, connection);
            adapter.Fill(ds);
            return ds.Tables[0];
        }
        public DataTable executeSelectSP()
        {
            DataSet ds = new DataSet();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(ds);
            return ds.Tables[0];
        }
        public DataTable executeSelectSP(MySqlCommand command)
        {
            DataSet ds = new DataSet();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(ds);
            return ds.Tables[0];
        }
        public IDataReader executeQuery(string sqlString)
        {
            command.CommandText = sqlString;
            return command.ExecuteReader();
        }
        public DataTableCollection executeCollectSelectSP(MySqlCommand command)
        {
            DataSet ds = new DataSet();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(ds);
            return ds.Tables;
        }
        public DataTableCollection executeCollectSelectSP()
        {
            DataSet ds = new DataSet();
            adapter = new MySqlDataAdapter(command);
            adapter.Fill(ds);
            return ds.Tables;
        }

        public void executeNonQuery(string sqlString)
        {
            command.CommandText = sqlString;
            command.ExecuteNonQuery();
        }
        public object executeSPScalar()
        {
            return command.ExecuteScalar();
        }
        public object executeSPScalar(MySqlCommand command)
        {
            return command.ExecuteScalar();
        }
        public void excuteSPNonQuery()
        {
            command.ExecuteNonQuery();
        }
        public object executeScalar(string sqlString)
        {
            command.CommandText = sqlString;
            return command.ExecuteScalar();
        }
        public object executeStoreProcedure()
        {
            return command.ExecuteScalar();
        }

        /// <summary>
        /// Chuyển một dòng dữ liệu thành một đối tượng tương ứng với lớp kế thừa
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        protected virtual object GetDataFromDataRow(DataTable dt, int i)
        {

            return null;
        }
        protected ArrayList ConvertDataSetToArrayList(DataSet dataset)
        {
            ArrayList arr = new ArrayList();
            DataTable dt = dataset.Tables[0];
            int i, n = dt.Rows.Count;
            for (i = 0; i < n; i++)
            {
                object hs = GetDataFromDataRow(dt, i);
                arr.Add(hs);

            }
            return arr;
        }
        protected ArrayList ConvertDataTableToArrayList(DataTable dt)
        {
            ArrayList arr = new ArrayList();
            int i, n = dt.Rows.Count;
            for (i = 0; i < n; i++)
            {
                object hs = GetDataFromDataRow(dt, i);
                arr.Add(hs);
            }
            return arr;
        }
        protected Object[] ConvertDataSetToArray(DataSet dataset)
        {
            DataTable dt = dataset.Tables[0];
            Object[] arr = new Object[dt.Rows.Count];
            int i, n = dt.Rows.Count;
            for (i = 0; i < n; i++)
            {
                object hs = GetDataFromDataRow(dt, i);
                arr[i] = hs;
            }
            return arr;
        }
        protected Object[] ConvertDataTableToArray(DataTable dt)
        {
            Object[] arr = new Object[dt.Rows.Count];
            int i, n = dt.Rows.Count;
            for (i = 0; i < n; i++)
            {
                object hs = GetDataFromDataRow(dt, i);
                arr[i] = hs;
            }
            return arr;
        }
        public DateTime GetDateSys()
        {
            connect();
            object date = executeScalar("sp_GetSysDate");
            disconnect();
            return (DateTime)date;
        }

        public bool TestConnect(string connectString)
        {
            MySqlConnection test = new MySqlConnection(connectString);
            try
            {
                test.Open();
                test.Close();
                return true;
            }
            catch { }
            return false;
        }

        public DataTable executeSelectQuery(string sqlString, string strConnection)
        {
            MySqlConnection test = new MySqlConnection(strConnection);
            try
            {
                DataSet ds = new DataSet();
                adapter = new MySqlDataAdapter(sqlString, test);
                adapter.Fill(ds);
                return ds.Tables[0];
            }
            catch { }
            return null;
        }

        public DataTable GetDataBases(string Server, string UserName, string Pass)
        {
            MySqlConnection conn = new MySqlConnection();
            try
            {
                //Server=.;Database=SiteCamera;uid=sa;pwd=qawsed;Connection Lifetime=100;Connect Timeout=500
                string connectString = "Server=" + Server + "; uid=" + UserName + ";pwd=" + Pass + "; Connection Lifetime=100;Connect Timeout=500";
                conn.ConnectionString = connectString;
                conn.Open();
                DataTable dt = conn.GetSchema("Databases");
                conn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                if (conn != null)
                    conn.Close();
                return null;
            }
        }

        public object executeScalar(string sqlString, string connectString)
        {
            MySqlConnection conn = new MySqlConnection();
            try
            {
                //Server=.;Database=SiteCamera;uid=sa;pwd=qawsed;Connection Lifetime=100;Connect Timeout=500                
                conn.ConnectionString = connectString;
                conn.Open();
                MySqlCommand command = new MySqlCommand(sqlString, conn);
                object result = command.ExecuteScalar();
                conn.Close();
                return result;
            }
            catch (Exception ex)
            {
                if (conn != null)
                    conn.Close();
                return null;
            }
        }
    }
}
