using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Minutas2
{
    public partial class VentaContadoParticulares : Form
    {
        public VentaContadoParticulares()
        {
            InitializeComponent();
        }

        private void btnvalidar_Click(object sender, EventArgs e)
        {
            object objMiss = System.Reflection.Missing.Value;
            Word.Application objword = new Word.Application();
            string ruta = Application.StartupPath + @"C:\Users\User\Desktop\Nueva carpeta\MINUTAS\VENTAS\Venta De Contado Entre Particulares.docx";
            object parametro = ruta;
            object numero_escritura = "numeroEP";
            Word.Document ObjDoc = objword.Documents.Open(parametro,objMiss);
            Word.Range num = ObjDoc.Bookmarks.get_Item(ref numero_escritura).Range;
            num.Text=txtnumescritura.Text;
            object rango1 = num;
            ObjDoc.Bookmarks.Add("numero_escritura", ref rango1);
            objword.Visible = true;
            
        }
    }
}
