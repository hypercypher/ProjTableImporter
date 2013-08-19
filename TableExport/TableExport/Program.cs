using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.IO;


/* TableImporter v0.1 - CPL 2013.08 
 * This creates an oledb connection to a MSAccess DB to create a table from a text file.
 * It builds the insertstatement dynamically when it parses through the file and appends it to a stringbuilder object.
 * Caveats: Currently cannot handle rows with missing field values.
 * Also, current version inserts one row at at time and is very time consuming. (Open connection, insert, close, repeat).
 * To Add: oledbconnection exception handler to handle inserting a row with missing fields.
 */
namespace TableImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            // TESTING lolcommits
            Console.WriteLine("Enter file name: ");
            string filename = Console.ReadLine().ToString();
            FileInfo fi = new FileInfo(string.Format("C:\\users\\hycy_tabby\\Databases\\{0}.txt",filename));
            StreamReader srdr = fi.OpenText();

            // open MSAccess db connection
            string connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\hycy_tabby\\Databases\\ImportAccessExample.accdb; Persist Security Info = False";
            OleDbConnection conn = new OleDbConnection(connString);

            string[] columnCollection;
            string[] rowCollection;
 
            string tableName = "table4";
            StringBuilder sbCreateTable = new StringBuilder();
            StringBuilder sbInsertTable = new StringBuilder();
            StringBuilder sbInsertTableValues = new StringBuilder();
            StringBuilder sbInsertStatement = new StringBuilder();

            // create new table with headers defined by the first row
            // generate insert statement
            string fullText = srdr.ReadToEnd();
            rowCollection = fullText.Split('\n');
            string[] header = rowCollection[0].Split('\t');

            sbCreateTable.Append(" CREATE TABLE " + tableName + " (");
            sbInsertTable.Append(" INSERT INTO " + tableName + " (");
                
            foreach (var field in header)
	        {   // remove return lines
                if (field.Contains('\r'))
                {
                    sbCreateTable.Append('[' + field.Substring(0, field.Length - 1) + ']' + " varchar(40) NULL,");
                    sbInsertTable.Append(field.Substring(0, field.Length)); // \r is removed before this line is executed.. interesting..
                    break;
                }
                // bracketed to handle headers named with reserved words
                sbCreateTable.Append('[' + field + ']' + " varchar(40) NULL,");
                sbInsertTable.Append('[' + field + ']' + ","); 
	        }
            sbCreateTable.Remove(sbCreateTable.Length - 1, 1);  // remove extra comma and return line
            sbInsertTable.Remove(sbInsertTable.Length - 1, 1);  // remove extra comma and return line
            sbCreateTable.Append(");");
            sbInsertTable.Append(") VALUES (");

            // create table
            executeNonQuery(conn, sbCreateTable.ToString());

            // generate insert statement 
            rowCollection = fullText.Split('\n');
            // parsing through each row
            foreach (string row in rowCollection)
            {
                columnCollection = row.Split('\t');
                // and each column of each of that row
                foreach (string column in columnCollection)
                {
                    sbInsertTableValues.Append("'" + column + "'" + ",");
                }
                sbInsertTableValues.Remove(sbInsertTableValues.Length - 1, 1);  // remove extra comma and return line

                sbInsertStatement.Append(sbInsertTable.ToString());
                sbInsertStatement.Append(sbInsertTableValues.ToString());
                sbInsertStatement.Append(");");   // close off insert statement
                
                // remove first and last curly brackets
                string temp = sbInsertStatement.ToString().Substring(1, sbInsertStatement.Length - 2);
                string temp2 = temp.Substring(0, temp.Length - 3);
                string temp3 = temp2 + "')";

                executeNonQuery(conn, temp3);   // see if we can insert multiple without closing.. nope

                // end transaction
                sbInsertTableValues.Clear();
                sbInsertStatement.Clear();
                
            }
            
            // exit gracefully
            conn.Close();
        }
        
        public static void executeNonQuery(OleDbConnection conn, string query)
        {
            conn.Open();
            OleDbCommand comm = new OleDbCommand(query, conn);
            try
            {
                comm.ExecuteNonQuery();
                Console.WriteLine(query);
            }
            catch (OleDbException exp)
            {
                // handle exception
            }
            conn.Close();
        }

        public static void readFromDB(OleDbConnection conn)
        {
            string testSelect = " SELECT * FROM ImportAccess";
            OleDbCommand comm = new OleDbCommand(testSelect, conn);
            OleDbDataReader dr = comm.ExecuteReader();
            object[] buffer = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

            StringBuilder sb = new StringBuilder();
            while (dr.Read())// populate buffer with values from db
            {
                int value = dr.GetValues(buffer);

                for (int i = 0; i < buffer.Length; i++)
                {
                    sb.Append(buffer[i] + ",");
                }
                sb.Remove(buffer.Length - 1, 1);
            }
            ConsoleKeyInfo cki = new ConsoleKeyInfo();
            while (true)    // keep console window open
            {
                Console.WriteLine("Returned values " + sb.ToString());
                cki = Console.ReadKey(true);
                if (cki.Key == ConsoleKey.X)
                {
                    break;
                }
            }
            dr.Close();
        }       

    }
}
