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

namespace Excel_12
{
    public partial class Form1 : Form
    {
        classic class1 = new classic();
        open class2 = new open();
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
            class2.Openbutton();
        }
    }
}
