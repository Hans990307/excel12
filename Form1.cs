using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Excel_12
{
    public partial class Form1 : Form
    {
        classic class1 = new classic();
        open Browse = new open();
        public Form1()
        {
            InitializeComponent();
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
           class1.nameSbv();
        }

        private void button2_Click(object sender, EventArgs e)
        {
   
            
            string filepath;
            string fileext;


            OpenFileDialog File1 = new OpenFileDialog(); //open dialog to choose the file
            if (File1.ShowDialog() == System.Windows.Forms.DialogResult.OK)//if user select a file
            {
                filepath = File1.FileName;//get the path of the file
                fileext = Path.GetExtension(filepath);//get the file extension
                if (fileext.CompareTo(".xls") == 0 || fileext.CompareTo(".xlxs") == 0)
                {
                    try
                    {
                        DataTable Data1 = new DataTable();//create datatable
                        Data1 = Browse.Openbutton(filepath, fileext);//call browse function
                        dataGridView1.Visible = true;//popout datagridview if true
                        dataGridView1.DataSource = Data1;//connect datatable with datagridview

                    }
                    catch
                    {

                        MessageBox.Show("Error");//show message if got error;


                    }
                }

            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
