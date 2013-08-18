using System;
using System.IO;
using System.Data.IO;
public class Class1
{
	public Class1()
	{

        string test = "INSERT INTO ImportAccess(Time,Source,Condition,Action,Level,Description,Value,Units,Operator) VALUES ('1','2','3','4','5','6','7','8','9');INSERT INTO ImportAccess(Time,Source,Condition,Action,Level,Description,Value,Units,Operator) VALUES ('1','2','3','4','5','6','7','8','9');";
        string connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\hycy_tabby\\Databases\\ImportAccessExample.accdb; Persist Security Info = False";
        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();
        executeNonQuery(conn, test);
	}

    public static void executeNonQuery(OleDbConnection conn, string insertStatement)
    {
        OleDbCommand comm = new OleDbCommand(insertStatement, conn);
        comm.ExecuteNonQuery();
    }


}
