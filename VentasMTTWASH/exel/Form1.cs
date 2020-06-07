using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using xls = Microsoft.Office.Interop.Excel;

namespace exel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        xls.Application x = new xls.Application();
        int i = 5;
        private void button1_Click(object sender, EventArgs e)
        {
            x.ActiveWorkbook.Worksheets[1].Cells(i, 1).Value = textBox1.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 2).Value = textBox2.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 3).Value = textBox3.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 4).Value = textBox4.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 5).Value = textBox5.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 6).Value = textBox6.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 7).Value = textBox7.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 8).Value = textBox8.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 9).Value = textBox9.Text;
            x.ActiveWorkbook.Worksheets[1].Cells(i, 10).Value = textBox10.Text;
           
            i++;
            MessageBox.Show("Datos guardados en el archivo de exel");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            x.Workbooks.Open(Application.StartupPath + @"\Datos.xlsx");
            x.Visible = true;
            while (x.ActiveWorkbook.ActiveSheet.Cells(i,1).Value != null)
            {
                i++;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            int j = 5;
            while (x.ActiveWorkbook.ActiveSheet.Cells(j, 1).Value != null)
            {
               
                string Tipo = x.ActiveWorkbook.ActiveSheet.Cells[j, 1].Value.ToString();
                string Tipodeservico = x.ActiveWorkbook.ActiveSheet.Cells[j, 2].Value.ToString();
                string Descripciondelservicio = x.ActiveWorkbook.ActiveSheet.Cells[j, 3].Value.ToString();
                string Preciodelservicio = x.ActiveWorkbook.ActiveSheet.Cells[j, 4].Value.ToString();
                string Cantidad = x.ActiveWorkbook.ActiveSheet.Cells[j, 5].Value.ToString();
                string Totaldelservicio = x.ActiveWorkbook.ActiveSheet.Cells[j, 6].Value.ToString();
                string Tipodepago = x.ActiveWorkbook.ActiveSheet.Cells[j, 7].Value.ToString();
                string Factura= x.ActiveWorkbook.ActiveSheet.Cells[j, 8].Value.ToString();
                string Entregado= x.ActiveWorkbook.ActiveSheet.Cells[j, 9].Value.ToString();
                string Clienteconforme= x.ActiveWorkbook.ActiveSheet.Cells[j, 10].Value.ToString();
                
                ListViewItem datos = new ListViewItem(Tipo);
                datos.SubItems.Add(Tipodeservico);
                datos.SubItems.Add(Descripciondelservicio);
                datos.SubItems.Add(Preciodelservicio);
                datos.SubItems.Add(Cantidad);
                datos.SubItems.Add(Totaldelservicio);
                datos.SubItems.Add(Tipodepago);
                datos.SubItems.Add(Factura);
                datos.SubItems.Add(Entregado);
                datos.SubItems.Add(Clienteconforme);
                
                listView1.Items.Add(datos);
                j++;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
