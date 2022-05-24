using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Lecture33
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb ;
        Worksheet ws ;
        public Form1()
        {
            InitializeComponent();
            wb = Excel.Workbooks.Open(@"C:\Users\Kiran Ramzan\Desktop\MPA1.xlsx");
            ws = wb.ActiveSheet;
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            // MessageBox.Show(ws.Shapes.Count.ToString()); //pictures are shapes in this dll

            ws.Shapes.Item(2).Copy();
            pictureBox1.Image = Clipboard.GetImage();

            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= ws.Shapes.Count; i++)
            {
                PictureBox pb = new PictureBox();
                if (i%2==0)
                {
                    ws.Shapes.Item(i).Copy();
                    pb.Image = Clipboard.GetImage();
                    flowLayoutPanel1.Controls.Add(pb);
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
            //string cell1 = ws.Cells[4, 2].Value.ToString();
            //lbl.Text = cell1;
            //flowLayoutPanel1.Controls.Add(lbl);
            //wb.Save();
            //wb.Close();
            for (int i = 1; i < ws.Cells[i, 2].Count; i++)
            {
                if (ws.Cells.Value(i).Copy()==true)
                {
                   System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
                    //ws.Cells.Value(i).Copy();
                    lbl.Text = Clipboard.GetText();
                    flowLayoutPanel1.Controls.Add(lbl);
                    wb.Save();
                    wb.Close();
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {     
           // string cell1 = ws.Cells[2,2].Value.ToString();
            for (int i = 1; i < ws.Cells[i,4].Count; i++)
            {
                System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
                if (i % 2 != 0)
                {
                    ws.Cells.Value(i).Copy();
                    lbl.Text = Clipboard.GetText();
                    flowLayoutPanel1.Controls.Add(lbl);
                    wb.Save();
                    wb.Close();
                }
            }
        }
    }
}
