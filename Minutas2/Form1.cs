using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace Minutas2
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            cmbMinutas.Items.Add("promesa de compraventa");
            cmbMinutas.Items.Add("venta");
            cmbMinutas.Items.Add("Poder general");

            cmbMinutas.SelectedIndex = 0;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (cmbMinutas.SelectedItem == null)
            {
                
                MessageBox.Show("Selecciona un formulario antes de continuar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string selectedForm = cmbMinutas.SelectedItem.ToString();

                
                switch (selectedForm)
                {
                    case "promesa de compraventa":
                        Form3 form3 = new Form3();
                        form3.Show();
                        break;

                    case "venta":
                        break;
                    case "Poder general":
                        Form2 form2 = new Form2();
                        
                        // Ruta del documento Word
                        string ruta = @"C:\Users\User\Desktop\minutas\404846-PODER GENERAL-MANUELA PATIÑO ARIAS - JHONY.docx";

                        // Verificar si el archivo existe
                        if (System.IO.File.Exists(ruta))
                        {
                            // Crear una instancia de la aplicación Word
                            Word.Application wordApp = new Word.Application();
                            Word.Document wordDoc = wordApp.Documents.Open(ruta);

                            // Obtener el texto del documento
                            string contenido = wordDoc.Content.Text;

                            // Cerrar el documento Word y la aplicación Word
                            wordDoc.Close();
                            wordApp.Quit();

                            // Asignar el contenido al RichTextBox
                            
                            form2.richTextBox1.Text = contenido;
                            
                            form2.Show();
                            
                            
                        }
                        else
                        {
                            MessageBox.Show("El archivo no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        break;

                    

                    default:
                        break;

                }// cierra case

            }// cierra else


            


        }//cierra metodo principañ
    }
}