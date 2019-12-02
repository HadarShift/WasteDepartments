using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

/// <summary>
/// Summary description for DbService
/// </summary>
public class DbServiceSQL
{
    OleDbTransaction tran;
    OleDbCommand cmd;
    //SqlConnection conn = new SqlConnection("Data Source=ALSQL;Initial Catalog=Production;Integrated Security=True;Uid=albi;Pwd=Al5342");
    //OleDbConnection conn = new OleDbConnection("Driver={SQL Server Native Client 11.0};Server=ALSQL\\ALLIANCE;Database=Production;Uid=albi;Pwd=Al5342");
    //SqlConnection conn = new SqlConnection("Provider=SQLNCLI11;Server=ALSQL\\ALLIANCE;Database=Production;Uid=albi;Pwd=Al5342");
    OleDbConnection conn = new OleDbConnection("Provider=SQLNCLI11;Server=ALFSQL;Database=PRODUCTION;Trusted_Connection = yes;Uid=albi;Pwd=Al5342");

    OleDbDataAdapter adp;
    public bool transactional = false;

    public DbServiceSQL()
    {
        adp = new OleDbDataAdapter();
        //conn = new SqlConnection("Data Source=ALSQL;Initial Catalog=Production;Integrated Security=True;Uid=albi;Pwd=Al5342");
        //conn = new SqlConnection("Driver={SQL Server Native Client 11.0};Server=ALSQL\\ALLIANCE;Database=Production;Uid=albi;Pwd=Al5342");
        //conn = new OleDbConnection("Data Source=ALSQL;Initial Catalog=Production;Integrated Security=True;Uid=albi;Pwd=Al5342");
        //
        // TODO: Add constructor logic here
        //
    }

    /// <method>
    /// Open Database Connection if Closed or Broken
    /// </method>
    private OleDbConnection openConnection()
    {
        if (conn.State == ConnectionState.Closed || conn.State == ConnectionState.Broken)
        {
            conn.Open();
        }
        return conn;
    }
    
   public DataTable executeSelectQueryNoParam(String _query)
        {
            DataTable dataTable = new DataTable();
            DataSet ds = new DataSet();
            using (OleDbConnection con = new OleDbConnection("Provider=SQLNCLI11;Server=ALSQL;Database=PRODUCTION;Trusted_Connection = yes;Uid=albi;Pwd=Al5342"))
            {
                con.Open();
                using (OleDbCommand myCommand = new OleDbCommand(_query,con))
                {
                    try
                    {
                    //myCommand.Connection = openConnection();
                    //myCommand.CommandText = _query;
                    //myCommand.CommandTimeout = 1000;
                    //myCommand.Parameters.AddRange(sqlParameter);
                    myCommand.ExecuteNonQuery();
                    adp.SelectCommand = myCommand;
                    adp.Fill(ds);
                    dataTable = ds.Tables[0];
                    }
                    catch (SqlException e)
                    {
                       //MessageBox.Show("Error - Connection.executeSelectQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
                        MessageBox.Show(e.Message);
                        return null;
                    }
                    finally
                    {
                        conn.Close();
                    }
                }
            }
            return dataTable;
        }
    


    public DataSet GetDataSetByQuery(string sqlQuery, CommandType cmdType = CommandType.Text, params SqlParameter[] parametersArray)
    {
        cmd = new OleDbCommand(sqlQuery, conn);
        cmd.CommandType = cmdType;
        DataSet ds = new DataSet();
        adp = new OleDbDataAdapter(cmd);

        foreach (SqlParameter s in parametersArray)
        {
            cmd.Parameters.AddWithValue(s.ParameterName, s.Value);

        }

        try
        {
            adp.Fill(ds);
        }
        catch (Exception e)
        {
            //do something with the error
            ds = null;
        }

        return ds;
    }

    public int GetScalarByQuery(string sqlQuery, CommandType cmdType = CommandType.Text, params SqlParameter[] parametersArray)
    {
        cmd = new OleDbCommand(sqlQuery, conn);
        cmd.CommandType = cmdType;
        int res = 0;

        string id = "0";
        foreach (SqlParameter s in parametersArray)
        {
            cmd.Parameters.AddWithValue(s.ParameterName, s.Value);
        }

        try
        {
            conn.Open();
            id = cmd.ExecuteScalar().ToString();
            res = Convert.ToInt32(id);



        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
            //do something with the error
        }
        finally
        {
            conn.Close();
        }



        return res;
    }


    public int ExecuteQuery(string sqlQuery, CommandType cmdType = CommandType.Text, params SqlParameter[] parametersArray)
    {
        int row_affected = 0;
        using (OleDbConnection conn = new OleDbConnection("Provider=SQLNCLI11;Server=ALSQL;Database=PRODUCTION;Trusted_Connection = yes;Uid=albi;Pwd=Al5342"))
        {

            conn.Open();
            tran = conn.BeginTransaction();

            cmd = new OleDbCommand(sqlQuery, conn, tran);
            cmd.CommandType = cmdType;

            foreach (SqlParameter s in parametersArray)
            {
                cmd.Parameters.AddWithValue(s.ParameterName, s.Value);
            }

            try
            {
                row_affected = cmd.ExecuteNonQuery();
                tran.Commit();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                tran.Rollback();
            }
            //finally
            //{
            //    con.Close();
            //}

        }

        return row_affected;
    }


    public object GetObjectScalarByQuery(string sqlQuery, CommandType cmdType = CommandType.Text, params SqlParameter[] parametersArray)
    {
        cmd = new OleDbCommand(sqlQuery, conn);
        cmd.CommandType = cmdType;
        object res = "0";

        foreach (SqlParameter s in parametersArray)
        {
            cmd.Parameters.AddWithValue(s.ParameterName, s.Value);
        }

        try
        {
            conn.Open();
            res = cmd.ExecuteScalar();
        }
        catch
        {
            //do something with the error
        }
        finally
        {
            conn.Close();
        }

        return res;
    }


    public void executeSelectQueryForDelete(String _query)
    {
        DataTable dataTable = new DataTable();
        DataSet ds = new DataSet();
        using (OleDbConnection con = new OleDbConnection("Provider=SQLNCLI11;Server=ALSQL;Database=PRODUCTION;Trusted_Connection = yes;Uid=albi;Pwd=Al5342"))
        {
            con.Open();
            using (OleDbCommand myCommand = new OleDbCommand(_query, con))
            {
                try
                {
                    //myCommand.Connection = openConnection();
                    //myCommand.CommandText = _query;
                    //myCommand.CommandTimeout = 1000;
                    //myCommand.Parameters.AddRange(sqlParameter);
                    myCommand.ExecuteNonQuery();
                    adp.SelectCommand = myCommand;
                    adp.Fill(ds);
                    
                }
                catch (SqlException e)
                {
                    //MessageBox.Show("Error - Connection.executeSelectQuery - Query: " + _query + " \nException: " + e.StackTrace.ToString());
                    MessageBox.Show(e.Message);
                    
                }
                finally
                {
                    conn.Close();
                }
            }
        }
    }

}


