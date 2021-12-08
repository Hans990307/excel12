using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace Excel_12
{
    public class open
    {
        public DataTable Openbutton(string Filename, string FileExtension) {

            string constr;
            string SelectF = "Select * from [$Sheet1A1:end]"; //select data inside the excel



            if (FileExtension.CompareTo(".xls") == 0) //if compare extension is the same
                constr = @"Provider=Microsoft.JET.OLEDB.4.0;Data Source=" +
                            Filename +
                            ";Extended Properties='Excel 8.0;HDR=YES;';"; //if the extension files is ".xls" this will be the connector path
            else
                constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            Filename +
                            ";Extended Properties='Excel 8.0;HDR=YES;';"; //if the extension files is ".xlxs" this will be the connector path

            DataTable Data1 = new DataTable(); //create a new data table
            OleDbConnection con = new OleDbConnection(constr);//Represents an open connection to a data source.
            OleDbDataAdapter dapter = new OleDbDataAdapter(SelectF, con); //Represents a set of data commands and a database connection that are used to fill the DataSet and update the data source.
            dapter.Fill(Data1);
            return Data1;


        }
    }
}
