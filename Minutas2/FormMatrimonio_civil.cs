using Word=Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Humanizer;
using System.Globalization;

namespace Minutas2
{
    public partial class formMatrimonio_civil : Form
    {
        public formMatrimonio_civil()
        {
          

            InitializeComponent();
            cmbnotario_encargado.Items.Add("JORGE MANRIQUE ANDRADE (T)");
            cmbnotario_encargado.Items.Add("CLAUDIA MARCELA GRANADA (E)");
            cmbnotario_encargado.Items.Add("MARCELA PATIÑO PEÑA (E)");

            cmbtitular.Items.Add("Titular");
            cmbtitular.Items.Add("Encargado");
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnvalidar_Click(object sender, EventArgs e)
        {
           


            
        }

        private void btnvalidar_Click_1(object sender, EventArgs e)
        {
            try
            {
                string fecha_otorgamiento = dtfecha_otorgamiento.Text.ToUpper();


                string nombre_otorgante1 = txtnombre_otorgante1.Text.ToUpper();
                string cedula_otorgante1 = txt_ced_otorgante1.Text;
                string ciud_exp_otorgante1 = txt_ciud_otorgante1.Text.ToUpper();
                string lugar_naci_otorgante1 = txtlugar_naci_otorgante1.Text.ToUpper();
                string fecha_naci_otorgante1 = dtfecha_naci_otorgante1.Text.ToUpper();


                string nombre_otorgante2 = txtnombre_otorgante2.Text.ToUpper();
                string cedula_otorgante2 = txt_ced_otorgante2.Text;
                string ciud_expd_otorgante2 = txt_ciud_otorgante2.Text.ToUpper();
                string lugar_naci_otorgante2 = txtlugar_naci_otorgante2.Text.ToUpper();
                string fecha_naci_otorgante2 = dtfecha_naci_otorgante2.Text.ToUpper();


                string minuta;
                string num_hojas1 = txtnum_hojas1.Text.ToUpper();
                string num_hojas2 = txtnum_hojas2.Text.ToUpper();
                string indicativo_serial = txtindicativo_serial.Text;
                string recaudo = txtrecaudo.Text;
                string acto = txtacto.Text;
                string derechos = txtderechos.Text;
                string iva = txtiva.Text;
                string elaboro = txtelaboro.Text.ToUpper();
                string firmas = txtfirmas.Text.ToUpper();
                string notario_encargado = cmbnotario_encargado.SelectedItem.ToString();
                string titular = cmbtitular.SelectedItem.ToString();
                string num_escritura = txtnumescritura.Text;

                DateTime fecha = ObtenerFechaDesdeCadena(fecha_otorgamiento);
                string cadenaFechaFormateada = FormatearFecha(fecha);

                int numero_letras = int.Parse(num_escritura);
                string numero_letras2 = ConvertirNumeroAPalabras(numero_letras).ToUpper();

                minuta = "ESCRITURA PÚBLICA NÚMERO: "+num_escritura+"-------------------------\r\nFECHA: "+fecha_otorgamiento+"-----\r\nCLASE DE ACTO: "+acto+". ---------------------------------------------------------\r\nOTORGANTES: "+nombre_otorgante1+", IDENTIFICADO CON LA CÉDULA DE CIUDADANÍA NÚMERO "+cedula_otorgante1+" EXPEDIDA EN "+ciud_exp_otorgante1+", Y "+nombre_otorgante2+", IDENTIFICADA CON LA CÉDULA DE CIUDADANÍA NÚMERO "+cedula_otorgante2 +" EXPEDIDA EN \r\n+ "+ciud_expd_otorgante2+". -------------------NOTARÍA DE ORIGEN: NOTARÍA SEGUNDA DE MANIZALES. -------------------------\r\nEn el municipio de Manizales, capital del departamento de Caldas, República de Colombia, al "+cadenaFechaFormateada +", en el despacho de la NOTARIA SEGUNDA DEL CÍRCULO DE MANIZALES a cargo del Notario(a) "+titular+" "+notario_encargado +", --------------------------------Comparecieron, el señor(a) "+nombre_otorgante1+", mayor de edad, vecino de Manizales, identificado con la cédula de ciudadanía número "+cedula_otorgante1+" expedida en "+ciud_exp_otorgante1+", nacido en "+lugar_naci_otorgante1+", el día "+fecha_naci_otorgante1+" de nacionalidad Colombiana y quien en la presente escritura se llamará EL CONTRAYENTE; y la señora "+nombre_otorgante2+", mayor de edad, vecino(a) de Manizales, identificado(a) con la cédula de ciudadanía número "+cedula_otorgante2+" expedida en "+ciud_expd_otorgante2+"., nacida en "+lugar_naci_otorgante2+", el día "+fecha_naci_otorgante1+", de nacionalidad Colombiana y quien en la presente escritura se llamará LA CONTRAYENTE; hábiles para contratar y obligarse, y dijeron: PRIMERO: Que en su entero y cabal juicio, es su deseo contraer matrimonio civil de conformidad con las prescripciones contenidas en el Decreto 2668 del 26 de Diciembre de 1.988. SEGUNDO: Que para tal efecto presentaron solicitud escrita y sus anexos, ante este despacho, todo lo cual se protocoliza con este instrumento público. TERCERO: Que constituidos en Audiencia Pública, el(la) Suscrito(a) Notario(a) preguntó claramente a los contrayentes si mediante el presente contrato de matrimonio, sin apremios de ninguna naturaleza, se quieren unir libre y espontáneamente, con el fin de formar una familia, vivir juntos, guardarse fe, socorrerse, procrear y ayudarse mutuamente en todas las circunstancias de la vida, con la afirmación de que el amor deberá presidir las relaciones entre los dos seres que por ministerio de la Ley quedan unidos en legítimo matrimonio ante la comunidad, procurando, con toda discreción y ternura, corregirse recíprocamente sus defectos, practicar la tolerancia y proceder en todos los casos con generosidad, equidad y templanza, evitando que entre ellos como esposos se presenten agravios de palabra o de obra que por su naturaleza irremediable comprometen la estabilidad del matrimonio, como comunión permanente entre dos seres que acuerdan transitar juntos el camino de la vida, como serían los hijos que llegaren a tener, preguntas y postulados anteriores todos los cuales los contrayentes, habiendo escuchado muy atentamente la lectura de esta escritura, manifiestan a el(la) Suscrito(a) Notario(a), con voz clara y perceptible que la han entendido completamente y por ello la aceptan y cumplirán íntegramente dichos postulados.- Agregan los contrayentes que el amor que los ha determinado para acogerse al vínculo matrimonial establecido por la Ley y por la  sociedad civil para perpetuar la especie les servirá para que en el transcurso de su vida estimulen una aproximación cada vez más estrecha entre  ellos  como  marido  y  mujer, para  así  entregarse por  entero  el  uno  al otro para la formación de la familia, con el pleno sentido de las responsabilidades que adquieren entre sí, para con sus  descendientes y frente a la comunidad a la cual pertenecen. Cada cual aportará su contingente, según las necesidades de la familia, para  constituirse en elementos de progreso ante la sociedad a la cual deberán  entregar, en el futuro, hijos y ciudadanos formados y educados en una atmósfera propicia para ser útiles a la familia, a la sociedad y a la patria; procurarán, en todo momento, que lo que ambos desearon al unirse en matrimonio no vaya a desmentirse  por  duras  que  sean  las  circunstancias que se les presente en el transcurso de su vida matrimonial.- CUARTO: Manifiestan los contrayentes al(la) Suscrito(a) Notario(a) que no tienen impedimento alguno para contraer matrimonio y que entre ellos no existe parentesco que pueda obstaculizarlo. – QUINTO: Que, en consecuencia, a partir de hoy los contrayentes se consideran unidos en legítimo matrimonio ante la sociedad y frente a las Leyes colombianas y aceptan cumplir fielmente los deberes y las obligaciones recíprocas que tal acto matrimonial trae consigo, respetándose los derechos naturales y civiles que cada uno tiene como persona humana, de conformidad con los preceptos establecidos en nuestras mencionadas leyes colombianas. La presente escritura de matrimonio fue leída personalmente por el(la) Suscrito(a) Notario(a) a los contrayentes, tal y conforme lo dispone la Ley, en clara y viva voz, la encontraron conforme a su voluntad, la aprobaron y la firman conmigo al(la) Suscrito(a) Notario(a) que de todo lo anterior doy fe. Los contrayentes presentaron los Siguientes documentos los cuales se protocolizan con el presente instrumento público: 1) Solicitud debidamente Autenticada para contraer matrimonio civil, 2) Fotocopias de las cédulas de ciudadanía de los contrayentes, 3) Fotocopias auténticas de los registros civiles de nacimiento de los contrayentes-----El presente matrimonio queda inscrito al indicativo serial número:\r\n "+ indicativo_serial+". --------- Así se firma en hojas de papel de seguridad notarial números: "+num_hojas1+"- "+num_hojas2+" -------------------------------------------------------------------------------------------\r\nDERECHOS: $"+derechos+" RECAUDOS $"+recaudo+" RESOLUCIÓN 00387 DE 23 DE ENERO DE 2023 DE LA SUPERINTENDENCIA DE NOTARIADO Y REGISTRO. IVA: $"+iva+" LEY 1819 DEL 29 DE DICIEMBRE DE 2016. RECEPCIONÓ: __________. ELABORÓ: "+elaboro+" FIRMAS: "+firmas+" CIERRE: _______. \"LO ESCRITO EN OTRO TIPO DE LETRA VALE\". ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\r\n\r\n\r\n\r\n\r\n\r\n\r\n_________________________________________\r\n"+nombre_otorgante1+"\r\nC.C. "+cedula_otorgante1+" EXPEDIDA EN "+ ciud_exp_otorgante1+".\r\nOCUPACIÓN: \r\nDIRECCIÓN: \r\nTELÉFONOS: \r\nCORREO ELECTRÓNICO: \r\n\r\n\r\n\r\nLA CONTRAYENTE:\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n_________________________________________\r\n"+nombre_otorgante2+"\r\nC.C. "+cedula_otorgante2+" EXPEDIDA EN "+ciud_expd_otorgante2+".\r\nOCUPACIÓN: \r\nDIRECCIÓN: \r\nTELÉFONOS:\r\nCORREO ELECTRÓNICO: \r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n"+notario_encargado+"\r\nNOTARIO(A) SEGUNDO(A) DEL CÍRCULO DE MANIZALES\r\n\r\n\r\n\r\n";

                richTextBox1.Text = minuta;


            }
            catch (Exception)
            {
                MessageBox.Show("error en los datos");
                throw;
            }
        }

        static DateTime ObtenerFechaDesdeCadena(string cadenaFechaOriginal)
        {
            // Puedes ajustar el formato según la cadena de entrada real
            return DateTime.ParseExact(cadenaFechaOriginal, "dddd, dd 'DE' MMMM 'DE' yyyy", CultureInfo.GetCultureInfo("es-ES"));
        }
       
        static string FormatearFecha(DateTime fecha)
        {
            DateTime fechaSeleccionada = DateTime.Now;
            // Obtener componentes de la fecha
            string dia = fecha.ToString("dd"); // Día en formato de dos dígitos
            string mes = fecha.ToString("MMMM", CultureInfo.GetCultureInfo("es-ES")); // Nombre del mes en español
            string año = ConvertirNumeroAPalabras(fechaSeleccionada.Year); // Año en formato de cuatro dígitos

            // Convertir el número del día a palabras
            string diaEnPalabras = ObtenerNumeroEnPalabras(dia);

            // Crear la cadena de fecha en el nuevo formato
            string resultado = $"{diaEnPalabras} ({dia}) DE {mes.ToUpper()} DE {año.ToUpper()} ({fecha.Year})";

            return resultado;
        }
        static string ObtenerNumeroEnPalabras(string numero)
        {
            {
                string[] unidades = { "CERO", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE" };
                string[] especiales = { "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE" };
                string[] decenas = { "CERO", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA" };

                int num = int.Parse(numero);

                if (num < 10)
                {
                    return unidades[num];
                }
                else if (num < 20)
                {
                    return especiales[num - 10];
                }
                else
                {
                    int decena = num / 10;
                    int unidad = num % 10;

                    if (unidad == 0)
                    {
                        return decenas[decena];
                    }
                    else
                    {
                        return $"{decenas[decena]} Y {unidades[unidad]}";
                    }
                }
            }
        }
        static string ConvertirNumeroAPalabras(int numero)
        {
            return numero.ToWords();
        }
        private void GuardarComoWord(string contenido)
        {
            try
            {
                // Crea una instancia de SaveFileDialog
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Documentos de Word (*.docx)|*.docx";
                saveFileDialog.Title = "Guardar como";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Crea una instancia de Word Application
                    Word.Application wordApp = new Word.Application();

                    // Crea un nuevo documento de Word
                    Word.Document doc = wordApp.Documents.Add();

                    // Agrega el contenido del RichTextBox al documento de Word
                    doc.Range().Text = contenido;

                    // Guarda el documento de Word en la ubicación elegida por el usuario
                    doc.SaveAs2(saveFileDialog.FileName, Word.WdSaveFormat.wdFormatDocumentDefault);

                    // Cierra Word
                    wordApp.Quit();

                    MessageBox.Show("Exportación y guardado exitosos en Word.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar y guardar en Word: " + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GuardarComoWord(richTextBox1.Text);
        }

        private void btnatras_Click(object sender, EventArgs e)
        {

            Main form1 = new Main();

            this.Close();
        }
    }///////cierra main
}
