using System.Text;
using System;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Sistema_de_mail_para_Bridgestone___Thalamus
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnFaltaClientes_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[,] datos = {
                { "Partner SKU Code", "COD-BRID", "Name" },
                { "Cliente 1", "", "" },
                { "Cliente 2", "", "" },
                { "Cliente 3", "", "" }
            };

            // Construir la tabla en formato de texto con tabulaciones
            StringBuilder plainText = new StringBuilder();
            StringBuilder html = new StringBuilder();

            // Encabezado del formato HTML para que Outlook lo reconozca
            html.AppendLine("Version:0.9");
            html.AppendLine("StartHTML:00000097");
            html.AppendLine("EndHTML:00000197");
            html.AppendLine("StartFragment:00000133");
            html.AppendLine("EndFragment:00000163");
            html.AppendLine("<html><body><table border='1' style='border-collapse:collapse;'>");
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
            html.AppendLine("</table></body></html>");

            // Crear un objeto DataObject y agregar texto en formato plano y HTML
            DataObject dataObject = new DataObject();
            dataObject.SetData(DataFormats.Text, plainText.ToString());
            dataObject.SetData(DataFormats.Html, html.ToString());

            // Copiar al portapapeles
            Clipboard.SetDataObject(dataObject);

            MessageBox.Show("Tabla copiada al portapapeles con formato", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile("C:\\Users\\Administrator\\Desktop\\stock.png");
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox4.Image = Image.FromFile("C:\\Users\\Administrator\\Desktop\\clientes.png");
            pictureBox4.SizeMode = PictureBoxSizeMode.Zoom;
        }

        private void BtnFaltaClientes_Click_1(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtClientesFal.Text))
            {
                Clipboard.SetText(

                    "Estimados, buenos días/tardes. \r\n\r\n" +
                    "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                    "Archivo de clientes:\r\n\r\n" +
                    "Detectamos que hay clientes que efectúan ventas que no están en el respectivo archivo.\r\n" +
                    "Favor de agregar los siguientes clientes:\r\n\r\n" +
                    txtClientesFal.Text + "\n\r\n" +
                    "Quedo al pendiente.\r\n" +
                    "Saludos.\r\n"

                    );
                MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("El cuadro de texto está vacío", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFiscal2.Checked)
            {
                panel3.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel3.Enabled = false; // Desactiva el Panel (se verá gris)
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

        private void checkErrorMail_CheckedChanged(object sender, EventArgs e)
        {
            if (checkErrorMail.Checked)
            {
                panel4.Enabled = true; // Desactiva el TextBox
            }
            else
            {
                panel4.Enabled = false;  // Activa el TextBox
            }
        }

        private void checkFaltaCli_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFaltaCli.Checked)
            {
                panel1.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel1.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkErrorFiscal_CheckedChanged(object sender, EventArgs e)
        {
            if (checkErrorFiscal.Checked)
            {
                panel2.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel2.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkEmailInvalido1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEmailInvalido1.Checked)
            {
                panel5.Enabled = true;  // Activa el Panel (puedes interactuar con él)
                button3.Enabled = true;
            }
            else
            {
                panel5.Enabled = false; // Desactiva el Panel (se verá gris)
                button3.Enabled = false;
            }
        }

        private void checkEmailInvalido2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEmailInvalido2.Checked)
            {
                textBox1.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                textBox1.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkEmailRepetido_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEmailRepetido.Checked)
            {
                panel6.Enabled = true;  // Activa el Panel (puedes interactuar con él)
                button2.Enabled = true;
            }
            else
            {
                panel6.Enabled = false; // Desactiva el Panel (se verá gris)
                button2.Enabled = false;
            }
        }

        private void checkEmailRepetido2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEmailRepetido2.Checked)
            {
                textBox2.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                textBox2.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkEmailInvalido2.Checked == false)
            {
                Clipboard.SetText(

                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\n" +
                "Detectamos que se ingresan correos no válidos en la columna “email”.\r\n\r\n" +
                "Favor de corregir y reenviar para poder reanudar con el procesamiento de los archivos\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"


                );
            }
            else if (checkEmailInvalido2.Checked == true)
            {
                Clipboard.SetText(
                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\n" +
                "Detectamos que se ingresan correos no válidos en la columna “email”.\r\n\r\n" +
                "Favor de corregir en las filas:\r\n\r\n" +
                textBox1.Text + "\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"
                );
            }
            MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (checkEmailRepetido2.Checked == false)
            {
                Clipboard.SetText(

                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\n" +
                "Detectamos que se ingresan correos duplicados en la columna “email”.\r\n\r\n" +
                "Favor de corregir y reenviar para poder reanudar con el procesamiento de los archivos\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"

                );
            }
            else if (checkEmailRepetido2.Checked == true)
            {
                Clipboard.SetText(
                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\nDetectamos que se ingresan correos duplicados en la columna “email”.\r\n\r\n" +
                "Favor de corregir en las filas:\r\n\r\n" +
                textBox2.Text + "\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"
                );

            }
            MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button4_Click(object sender, EventArgs e)

        {

            StringBuilder mensaje = new StringBuilder();

            mensaje.AppendLine("Estimados, buenos días/tardes.\r\n");
            mensaje.AppendLine("Sobre sus archivos les comento lo siguiente:\r\n");

            int contadorErrores = 0; // Contador para numerar los errores

            if (checkFaltaCli.Checked)
            {
                contadorErrores++;
                mensaje.AppendLine($"{contadorErrores}. Hay clientes que efectúan ventas que no están en el respectivo archivo.");
                mensaje.AppendLine("Favor de agregar los siguientes clientes:\r\n");
                mensaje.AppendLine(txtClientesFal.Text);
                mensaje.AppendLine("\r\n");
            }

            if (checkErrorFiscal.Checked)
            {
                contadorErrores++;
                mensaje.AppendLine($"{contadorErrores}. En ciertas celdas de la columna “fiscalId” se ingresan caracteres inválidos (guiones, puntos, espacios, etc).");
                mensaje.AppendLine("Recordar que el formato de la columna “fiscalId” es completamente numérico.");
                if (!string.IsNullOrWhiteSpace(txtFiscalId.Text))
                {
                    mensaje.AppendLine("Favor de corregir en las filas:\r\n");
                    mensaje.AppendLine(txtFiscalId.Text);
                }
                mensaje.AppendLine("\r\n");
            }

            if (checkErrorMail.Checked)
            {
                contadorErrores++;
                mensaje.AppendLine($"{contadorErrores}. Se ingresan correos no válidos en la columna “email”.");
                mensaje.AppendLine("Favor de corregir y reenviar para poder reanudar con el procesamiento de los archivos");
                if (!string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    mensaje.AppendLine("Favor de corregir en las filas:\r\n");
                    mensaje.AppendLine(textBox1.Text);
                }
                mensaje.AppendLine("\r\n");
            }

            if (checkEmailRepetido.Checked)
            {
                contadorErrores++;
                mensaje.AppendLine($"{contadorErrores}. Se ingresan correos duplicados en la columna “email”. Recordar que solo se puede ingresar un mail único");
                if (!string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    mensaje.AppendLine("Favor de corregir en las filas:\r\n");
                    mensaje.AppendLine(textBox2.Text);
                }
                mensaje.AppendLine("\r\n");
            }

            mensaje.AppendLine("Quedo al pendiente.");
            mensaje.AppendLine("Saludos.");

            // Copia el texto generado al portapapeles
            Clipboard.SetText(mensaje.ToString());

            MessageBox.Show("Texto combinado copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkFiscal2.Checked == false)
            {
                Clipboard.SetText(

                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\n" +
                "Detectamos que en ciertas celdas de la columna “fiscalId” se ingresan caracteres inválidos (guiones, puntos, espacios, etc), los cuales imposibilitan el procesamiento de la fila.\r\n" +
                "Recordar que el formato de la columna “fiscalId” es completamente numérica, cualquier carácter no numérico dará error.\r\n\r\n" +
                "Favor de corregir y reenviar para poder reanudar con el procesamiento de los archivos\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos.\r\n"


                );
            }
            else if (checkFiscal2.Checked == true)
            {
                Clipboard.SetText(
                "Estimados, buenos días/tardes. \r\n\r\n" +
                "Sobre sus archivos les comento lo siguiente:\r\n\r\n" +
                "Archivo de clientes:\r\n\r\n" +
                "Detectamos que en ciertas celdas de la columna “fiscalId” se ingresan caracteres inválidos (guiones, puntos, espacios, etc), los cuales imposibilitan el procesamiento de la fila.\r\n" +
                "Recordar que el formato de la columna “fiscalId” es completamente numérica, cualquier carácter no numérico dará error.\r\n\r\n" +
                "Favor de corregir en las filas:\r\n\r\n" +
                txtFiscalId.Text +
                "\r\n\r\n" +
                "Quedo al pendiente.\r\n" +
                "Saludos."
                );
            }
            MessageBox.Show("Texto copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                panel7.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel7.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                panel14.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel14.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkBox10_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                panel15.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                panel15.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xls",
                Title = "Seleccionar archivo Excel"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ProcesarExcel(openFileDialog.FileName);
            }
        }

        private void ProcesarExcel(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Suponemos que es la primera hoja
                int rowCount = worksheet.Dimension.Rows;
                List<string> facturasPromo = new List<string>();
                List<string> facturasAfiliados = new List<string>();

                for (int row = 2; row <= rowCount; row++) // Asumiendo que la primera fila son encabezados
                {
                    string tipoPromo = worksheet.Cells[row, 9].Text.Trim(); // Columna H
                    string tipoAfiliado = worksheet.Cells[row, 7].Text.Trim(); // Columna J
                    string factura = worksheet.Cells[row, 10].Text.Trim(); // Número de factura en la misma columna J

                    if (tipoPromo.IndexOf("promo", StringComparison.OrdinalIgnoreCase) >= 0 && tipoPromo.IndexOf("no promo", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        facturasPromo.Add(factura);
                    }

                    if (tipoAfiliado.IndexOf("afiliad", StringComparison.OrdinalIgnoreCase) >= 0 && !tipoAfiliado.Equals("No afiliado", StringComparison.OrdinalIgnoreCase))
                    {
                        facturasAfiliados.Add(factura);
                    }
                }

                // Tomar el 5% aleatorio de cada lista
                Random random = new Random();
                var seleccionPromo = facturasPromo.OrderBy(x => random.Next()).Take((int)Math.Ceiling(facturasPromo.Count * 0.05)).ToList();
                var seleccionAfiliados = facturasAfiliados.OrderBy(x => random.Next()).Take((int)Math.Ceiling(facturasAfiliados.Count * 0.05)).ToList();

                // Construir el mensaje
                StringBuilder mensaje = new StringBuilder();
                mensaje.AppendLine("Estimada, buenos días.");
                mensaje.AppendLine();
                mensaje.AppendLine("Me comunico, a pedido de Bridgestone y con el fin de realizar la auditoría de promociones, para solicitarle que suba al FTP, dentro de la carpeta Facturas/promociones/febrero, las siguientes facturas:\r\n");
                mensaje.AppendLine(string.Join("\r\n", seleccionPromo));
                mensaje.AppendLine();
                mensaje.AppendLine("----");
                mensaje.AppendLine();
                mensaje.AppendLine("También a pedido de Bridgestone y con el fin de realizar la auditoría de AFILIADOS, para solicitarle que suba al FTP, dentro de la carpeta Facturas/afiliados/febrero, las siguientes facturas:\r\n");
                mensaje.AppendLine(string.Join("\r\n", seleccionAfiliados));
                mensaje.AppendLine();
                mensaje.AppendLine("---");
                mensaje.AppendLine();
                mensaje.AppendLine("Quedo al pendiente.");
                mensaje.AppendLine("Saludos.");

                // Copiar al portapapeles
                Clipboard.SetText(mensaje.ToString());
                MessageBox.Show("Mensaje copiado al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
    


