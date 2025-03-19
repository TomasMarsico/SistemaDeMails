using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Sistema_de_mail_para_Bridgestone___Thalamus
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

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
                //panel3.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //panel3.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

       

        private void checkErrorMail_CheckedChanged(object sender, EventArgs e)
        {
            if (checkErrorMail.Checked)
            {
                //panel4.Enabled = true; // Desactiva el TextBox
            }
            else
            {
                // panel4.Enabled = false;  // Activa el TextBox
            }
        }

        private void checkFaltaCli_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFaltaCli.Checked)
            {
                //  panel1.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //  panel1.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkErrorFiscal_CheckedChanged(object sender, EventArgs e)
        {
            if (checkErrorFiscal.Checked)
            {
                //  panel2.Enabled = true;  // Activa el Panel (puedes interactuar con él)
            }
            else
            {
                //  panel2.Enabled = false; // Desactiva el Panel (se verá gris)
            }
        }

        private void checkEmailInvalido1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEmailInvalido1.Checked)
            {
                //  panel5.Enabled = true;  // Activa el Panel (puedes interactuar con él)
                button3.Enabled = true;
            }
            else
            {
                //  panel5.Enabled = false; // Desactiva el Panel (se verá gris)
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
                //  panel6.Enabled = true;  // Activa el Panel (puedes interactuar con él)
                button2.Enabled = true;
            }
            else
            {
                // panel6.Enabled = false; // Desactiva el Panel (se verá gris)
                button2.Enabled = false;
            }
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
            mensaje.AppendLine("Archivo de clientes:\r\n");

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
    }
}
