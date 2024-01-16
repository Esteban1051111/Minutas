using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;


namespace Minutas2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
                object ObjMiss = System.Reflection.Missing.Value;
                word.Application ObjWord = new word.Application();

                string ruta = @"C:\Users\User\Downloads\Modelo_de_Promesa_de_Compraventa_de_bien_Inmueble.doc";
                object parametro = ruta;
                object nombre1 = "nombre1";
                object cedula = "cedula";
                object ciudad = "ciudad";

                word.Document ObjDoc = ObjWord.Documents.Open(ref parametro, ref ObjMiss);
                word.Range non = ObjDoc.Bookmarks.get_Item(ref nombre1).Range;
                non.Text = textBox1.Text;

                word.Range ced = ObjDoc.Bookmarks.get_Item(ref cedula).Range;
                ced.Text = textBox2.Text;

                word.Range ciu = ObjDoc.Bookmarks.get_Item(ref ciudad).Range;
                ciu.Text = textBox3.Text;

                object rango1 = non;
                object rango2 = ced;
                object rango3 = ciu;

                ObjDoc.Bookmarks.Add("nombre", ref rango1);
                ObjDoc.Bookmarks.Add("cedula", ref rango2);
                ObjDoc.Bookmarks.Add("ciudad", ref rango3);
                richTextBox1.Text = ObjDoc.Content.Text;

            
            
            

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnatras_Click(object sender, EventArgs e)
        {
            Main form1 = new Main();

            this.Close();
            
            
        }
    }
}
