using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sistema_de_mail_para_Bridgestone___Thalamus
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
        private void btnFaltaSkuStock_Click(object sender, EventArgs e)
        {
            string mensajeTexto = "Estimados, buenos días/tardes.\r\n\r\n" +
                          "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                          "Archivo de stock:\r\n\r\n" +
                          "Detectamos que se ingresan códigos no coincidentes con los del archivo de equivalencias.\r\n" +
                          "Favor de completar el siguiente cuadro:\r\n\r\n";

            // Obtener los datos del TextBox
            string[] skuCodes = txtPartnerSKU.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            // Definir la estructura de la tabla
            string[,] datos = new string[skuCodes.Length + 1, 3];
            datos[0, 0] = "Partner SKU Code";
            datos[0, 1] = "COD-BRID";
            datos[0, 2] = "Name";

            for (int i = 0; i < skuCodes.Length; i++)
            {
                datos[i + 1, 0] = skuCodes[i];
                datos[i + 1, 1] = "";
                datos[i + 1, 2] = "";
            }

            // Construir la tabla en formato de texto con tabulaciones
            StringBuilder plainText = new StringBuilder();
            StringBuilder html = new StringBuilder();

            // Encabezado del formato HTML compatible con Outlook y Gmail
            html.AppendLine("Version:0.9");
            html.AppendLine("StartHTML:00000097");
            html.AppendLine("EndHTML:00000197");
            html.AppendLine("StartFragment:00000133");
            html.AppendLine("EndFragment:00000163");
            html.AppendLine("<html><body>");
            html.AppendLine($"<p>{mensajeTexto.Replace("\r\n", "<br>")}</p>");
            html.AppendLine("<table border='1' style='border-collapse:collapse;'>");
            html.AppendLine("<!--StartFragment-->");

            for (int i = 0; i < datos.GetLength(0); i++)
            {
                html.AppendLine("<tr>");
                for (int j = 0; j < datos.GetLength(1); j++)
                {
                    html.AppendLine($"<td style='padding:5px;'>{datos[i, j]}</td>");
                    plainText.Append(datos[i, j] + "\t");
                }
                plainText.AppendLine();
                html.AppendLine("</tr>");
            }

            html.AppendLine("<!--EndFragment-->");
            html.AppendLine("</table><br>");
            html.AppendLine("<p>Quedo al pendiente.<br>Saludos.</p>");
            html.AppendLine("</body></html>");

            // Crear un objeto DataObject y agregar el mensaje en formato plano y HTML
            DataObject dataObject = new DataObject();
            dataObject.SetData(DataFormats.Text, mensajeTexto + plainText.ToString());
            dataObject.SetData(DataFormats.Html, html.ToString());

            // Copiar al portapapeles
            Clipboard.SetDataObject(dataObject);

            MessageBox.Show("Mensaje con tabla copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        private void btnErrorCodSucursal_Click(object sender, EventArgs e)
        {
            if (checkBox5.Checked == false)
            {
                Clipboard.SetText(

                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de stock:\r\n\r\n" +
                "Detectamos que se ingresan códigos de sucursal que no fueron declarados anteriormente en el archivo de sucursales correspondiente.\r\n" +
                "Favor de corregir y reenviar.\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos."


                );
            }
            else if (checkBox5.Checked == true)
            {
                Clipboard.SetText(
                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de stock:\r\n\r\n" +
                "Detectamos que se ingresan códigos de sucursal que no fueron declarados anteriormente en el archivo de sucursales correspondiente.\r\n\r\n" +
                "Codigos declarados:\r\n\r\n" +
                textBox4.Text + "\r\n\r\n" +
                "Códigos ingresados en este archivo:\r\n\r\n" +
                textBox3.Text + "\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"
                );
            }
            MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                //    panel7.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //panel7.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                //panel14.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //panel14.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkBox10_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                //panel15.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //panel15.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }
        private void btnErrorStockDate_Click(object sender, EventArgs e)
        {

            {
                if (checkBox10.Checked == false)
                {
                    Clipboard.SetText(

                    "Estimados, buenos días/tardes. \r\n\r\n" +
                    "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                    "Archivo de stock:\r\n\r\n" +
                    "Detectamos que en la columna “Stock Date” se ingresan fechas erróneas o en formato incorrecto.\r\n" +
                    "Recordar que el formato que manejamos es \"fecha corta\", por tanto cualquier otro formato genera error en el sistema.\r\n" +
                    "Favor de corregir y reenviar.\r\n\r\n" +
                    "Quedo al pendiente.\r\n" +
                    "Saludos."


                    );
                }
                else if (checkBox10.Checked == true)
                {
                    Clipboard.SetText(
                    "Estimados, buenos días/tardes. \r\n\r\n" +
                    "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                    "Archivo de stock:\r\n\r\n" +
                    "Detectamos que en la columna “Stock Date” se ingresan fechas erróneas o en formato incorrecto.\r\n" +
                    "Recordar que el formato que manejamos es \"fecha corta\", por tanto cualquier otro formato genera error en el sistema.\r\n\r\n" +
                    "Favor de corregir en las filas:\r\n" +
                    textBox7.Text +
                    "\r\n\r\n" +
                    "Quedo al pendiente.\r\n" +
                    "Saludos."
                    );
                }
                MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnErrorCodSucursal_Click_1(object sender, EventArgs e)
        {

        }
    }
}
