using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
//using MySql.Data.MySqlClient;
using System.Data.SqlClient;
//using System.Data.OleDb.
using System.Net;
using System.Text.RegularExpressions;

namespace OrderTracker
{
   public partial class Settings
    {
        public SqlConnection conn = new SqlConnection();
        public SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        SqlDataAdapter da;

        DataSet ds;
        /// <summary>
        /// This function will get the connnection string from web.config
        /// </summary>
        /// <returns>connection string</returns>
        /// <remarks></remarks>
        public string getConnectionString()
        {
            string objConnectionString = "";
            //objConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\msstock.mdb;Jet OLEDB:Database Password=ms;"                   
            return objConnectionString;
        }



        /// <summary>
        /// This function will create a connection to database
        /// </summary>
        /// <remarks></remarks>
        public void createConnection()
        {
            string connstr = getConnectionString();
            conn = new SqlConnection(connstr);
            try
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();

            }
            catch (Exception ex)
            {
            }
        }




        /// <summary>
        /// This function will help to insert values to database.
        /// </summary>
        /// <param name="sqlStr">This is a sql string for insert to database</param>
        /// <returns>A integer will return.</returns>
        /// <remarks>If the function returns a non zero value then the values submitted succesfully.</remarks>
        public int InsertOrUpdateOrDeleteValueToDatabase(string sqlStr)
        {
            int Status = 0;

            try
            {
                createConnection();
                cmd = new SqlCommand(sqlStr, conn);
                cmd.CommandTimeout = 0;
                Status = cmd.ExecuteNonQuery();

                return Status;

            }
            catch (Exception ex)
            {
                return Status;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <returns>gd</returns>
        /// <remarks>sdfsdfsdfsd sdsdfsdf sdfsdfsd sdfsdf</remarks>
        public DataSet selectAllfromDatabaseAndReturnDataSet(string sql)
        {
            try
            {
                createConnection();
                cmd = new SqlCommand(sql, conn);
                da = new SqlDataAdapter(cmd);
                ds = new DataSet();
                da.Fill(ds, "default");

            }
            catch (Exception ex)
            {
                ds = null;
            }

            return ds;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql">sql select command</param>
        /// <returns>Mysqldatareader</returns>
        /// <remarks></remarks>
        public SqlDataReader selectFromDataBaseAndreturnDatareader(string sql)
        {
            try
            {
                createConnection();
                cmd = new SqlCommand(sql, conn);
                dr = cmd.ExecuteReader();

            }
            catch (Exception ex)
            {
                dr = null;
            }

            return dr;
        }
        /// <summary>
        /// insert apostropy in database
        /// </summary>
        /// <param name="text">string to be replaced</param>
        /// <returns>string</returns>
        /// <remarks></remarks>
        public object apostropy(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }
            else
            {
                text = text.Replace("'", "''");
                return text;
            }

        }
   
    }
}
