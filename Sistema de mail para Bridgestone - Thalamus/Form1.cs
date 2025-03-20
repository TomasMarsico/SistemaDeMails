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
        string tipoDeArch = "";
        private bool isDragging = false;
        private Point startPoint = new Point(0, 0);

        public Form1()
        {
            InitializeComponent();
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            panel3.Visible = true;

            // Simular una operación que toma tiempo
            await Task.Delay(750); // Simula una carga de 3 segundos

            // Ocultar la barra de carga
            panel3.Visible = false;

        }

        private async void selecFcsBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xls",
                Title = "Seleccionar archivo Excel"
            };
            panel3.Visible = true;

            // Simular una operación que toma tiempo
            await Task.Delay(650); // Simula una carga de 3 segundos

            // Ocultar la barra de carga
            panel3.Visible = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ProcesarFacturas(openFileDialog.FileName);
            }

        }

        private void ProcesarFacturas(string filePath)
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

        private async void selecArchBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xls",
                Title = "Seleccionar archivos Excel",
                Multiselect = true // Permitir selección múltiple
            };
            panel3.Visible = true;

            // Simular una operación que toma tiempo
            await Task.Delay(650); // Simula una carga de 3 segundos

            // Ocultar la barra de carga
            panel3.Visible = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StringBuilder mensajeCombinadoTexto = new StringBuilder();
                StringBuilder mensajeCombinadoHtml = new StringBuilder();

                // Encabezado común para HTML
                mensajeCombinadoHtml.AppendLine("Version:0.9");
                mensajeCombinadoHtml.AppendLine("StartHTML:00000097");
                mensajeCombinadoHtml.AppendLine("EndHTML:00000197");
                mensajeCombinadoHtml.AppendLine("StartFragment:00000133");
                mensajeCombinadoHtml.AppendLine("EndFragment:00000163");
                mensajeCombinadoHtml.AppendLine("<html><body>");

                foreach (string filePath in openFileDialog.FileNames)
                {
                    var (mensajeTexto, mensajeHtml) = ProcesarArchivoIndividual(filePath);

                    mensajeCombinadoTexto.AppendLine(mensajeTexto);
                    mensajeCombinadoTexto.AppendLine("---"); // Separador entre archivos

                    mensajeCombinadoHtml.AppendLine(mensajeHtml);
                    mensajeCombinadoHtml.AppendLine("<hr>"); // Separador entre archivos en HTML
                }

                // Cierre del HTML
                mensajeCombinadoHtml.AppendLine("</body></html>");

                // Copiar al portapapeles
                DataObject dataObject = new DataObject();
                dataObject.SetData(DataFormats.Text, mensajeCombinadoTexto.ToString());
                dataObject.SetData(DataFormats.Html, mensajeCombinadoHtml.ToString());
                Clipboard.SetDataObject(dataObject);

                MessageBox.Show("Mensajes combinados copiados al portapapeles", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private (string mensajeTexto, string mensajeHtml) ProcesarArchivoIndividual(string filePath)
        {
            StringBuilder mensajeTexto = new StringBuilder();
            StringBuilder mensajeHtml = new StringBuilder();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension?.Columns ?? 0; // Usar el operador null-conditional para evitar excepciones
                int rowCount = worksheet.Dimension?.Rows ?? 0;

                if (colCount == 0 || rowCount == 0)
                {
                    MessageBox.Show($"El archivo {Path.GetFileName(filePath)} está vacío o no tiene datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return (mensajeTexto.ToString(), mensajeHtml.ToString());
                }

                int lastColWithData = 0;

                // Identificar la última columna con datos en la primera fila (encabezados)
                for (int col = 1; col <= colCount; col++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cells[1, col].Text))
                    {
                        lastColWithData = col;
                    }
                }

                // Mostrar el valor de lastColWithData en un Label (para depuración)
                label2.Text = $"Última columna con datos: {lastColWithData}";

                if (lastColWithData == 0)
                {
                    MessageBox.Show($"No se encontraron columnas con datos en el archivo {Path.GetFileName(filePath)}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return (mensajeTexto.ToString(), mensajeHtml.ToString());
                }

                string statusHeader = worksheet.Cells[1, lastColWithData].Text.Trim().ToLower();

                if (statusHeader != "status")
                {
                    MessageBox.Show($"El archivo {Path.GetFileName(filePath)} no tiene la columna 'status' en la última columna con datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return (mensajeTexto.ToString(), mensajeHtml.ToString());
                }

                switch (lastColWithData)
                {
                    case 5:
                        if (worksheet.Cells[1, 5].Text.Trim().Equals("status", StringComparison.OrdinalIgnoreCase))
                        {
                            tipoDeArch = "Archivo de stock";
                            var (textoStock, htmlStock) = ProcesarStock(filePath);
                            mensajeTexto.AppendLine(textoStock);
                            mensajeHtml.AppendLine(htmlStock);
                        }
                        break;
                    case 11:
                        if (worksheet.Cells[1, 11].Text.Trim().Equals("status", StringComparison.OrdinalIgnoreCase))
                        {
                            tipoDeArch = "Archivo de ventas";
                            var (textoVentas, htmlVentas) = ProcesarVentas(filePath);
                            mensajeTexto.AppendLine(textoVentas);
                            mensajeHtml.AppendLine(htmlVentas);
                        }
                        break;
                    case 14:
                        if (worksheet.Cells[1, 14].Text.Trim().Equals("status", StringComparison.OrdinalIgnoreCase))
                        {
                            tipoDeArch = "Archivo de clientes";
                            var (textoClientes, htmlClientes) = ProcesarClientes(filePath);
                            mensajeTexto.AppendLine(textoClientes);
                            mensajeHtml.AppendLine(htmlClientes);
                        }
                        break;
                    default:
                        mensajeTexto.AppendLine($"El archivo {Path.GetFileName(filePath)} no es compatible.");
                        mensajeHtml.AppendLine($"<p>El archivo {Path.GetFileName(filePath)} no es compatible.</p>");
                        break;
                }
            }

            return (mensajeTexto.ToString(), mensajeHtml.ToString());
        }

        private (string mensajeTexto, string mensajeHtml) ProcesarStock(string filePath)
        {
            StringBuilder mensajeTexto = new StringBuilder();
            StringBuilder mensajeHtml = new StringBuilder();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                Dictionary<string, HashSet<string>> erroresSku = new Dictionary<string, HashSet<string>>();
                Dictionary<string, List<string>> otrosErrores = new Dictionary<string, List<string>>();

                for (int row = 2; row <= rowCount; row++)
                {
                    string mensajeErrorColumna = worksheet.Cells[row, 5].Text.Trim();
                    string skuCode = worksheet.Cells[row, 1].Text.Trim();
                    string quantity = worksheet.Cells[row, 3].Text.Trim();
                    string mensajeError = ObtenerMensajeErrorStock(mensajeErrorColumna, row, skuCode, quantity);

                    if (!string.IsNullOrEmpty(mensajeError))
                    {
                        if (mensajeError.StartsWith("Detectamos que se ingresan códigos"))
                        {
                            if (!erroresSku.ContainsKey(mensajeError))
                            {
                                erroresSku[mensajeError] = new HashSet<string>();
                            }
                            erroresSku[mensajeError].Add(skuCode);
                        }
                        else
                        {
                            if (!otrosErrores.ContainsKey(mensajeError))
                            {
                                otrosErrores[mensajeError] = new List<string>();
                            }
                            otrosErrores[mensajeError].Add(row.ToString());
                        }
                    }
                }

                // Construir mensaje de stock
                mensajeTexto.AppendLine($"Archivo de stock:");
                mensajeHtml.AppendLine($"<p><span style='text-decoration: underline;'>Archivo de stock:</span></p>");

                if (erroresSku.Count > 0 || otrosErrores.Count > 0)
                {
                    mensajeTexto.AppendLine("Se detectaron los siguientes errores en el archivo:");
                    mensajeHtml.AppendLine("<p>Se detectaron los siguientes errores en el archivo:</p>");

                    // Errores de SKU
                    if (erroresSku.Count > 0)
                    {
                        mensajeHtml.Append(ConstruirTablaSku(erroresSku.First().Value));
                        mensajeTexto.AppendLine("---");
                        mensajeHtml.AppendLine("<hr>");
                    }

                    // Otros errores
                    bool primerError = true;
                    if (otrosErrores.Count > 0)
                    {
                        foreach (var error in otrosErrores)
                        {
                            if (primerError)
                            {
                                mensajeTexto.AppendLine($"{error.Key}");
                                mensajeHtml.AppendLine($"<p>{error.Key.Replace(":", ".")}</p>");
                                primerError = false;
                            }
                            else
                            {
                                mensajeTexto.AppendLine("---");
                                mensajeHtml.AppendLine("<hr>");
                                mensajeTexto.AppendLine($"También {error.Key.ToLower().Replace("detectamos que", "detectamos que")}");
                                mensajeHtml.AppendLine($"<p>También {error.Key.Replace(":", ".").ToLower().Replace("detectamos que", "detectamos que")}</p>");
                            }

                            if (error.Key.StartsWith("Detectamos que se ingresan clientes"))
                            {
                                mensajeTexto.AppendLine("   Favor de agregar los siguientes clientes:");
                                mensajeTexto.AppendLine(string.Join(Environment.NewLine, error.Value.Distinct()));
                                mensajeHtml.AppendLine("   <p>Favor de agregar los siguientes clientes:</p>");
                                mensajeHtml.AppendLine($"   <p>{string.Join("<br>", error.Value.Distinct())}</p>");
                            }
                            else if (error.Key.StartsWith("Detectamos que se ingresan ventas con el formato de fecha"))
                            {
                                mensajeTexto.AppendLine("   Favor de corregir y reenviar.");
                                mensajeHtml.AppendLine("   <p>Favor de corregir y reenviar.</p>");
                            }
                            else if (error.Key.StartsWith("Detectamos que se ingresan comas"))
                            {
                                mensajeTexto.AppendLine("   Favor de corregir y reenviar.");
                                mensajeHtml.AppendLine("   <p>Favor de corregir y reenviar.</p>");
                            }
                            else
                            {
                                mensajeTexto.AppendLine("   Favor de corregir en las filas:");
                                mensajeTexto.AppendLine($"    {string.Join(", ", error.Value.Distinct())}");
                                mensajeHtml.AppendLine($"    <p>{string.Join(", ", error.Value.Distinct())}</p>");
                            }
                        }
                    }
                }
                else
                {
                    mensajeTexto.AppendLine("No se detectaron errores en el archivo.");
                    mensajeHtml.AppendLine("<p>No se detectaron errores en el archivo.</p>");
                }
            }

            return (mensajeTexto.ToString(), mensajeHtml.ToString());
        }

        private (string mensajeTexto, string mensajeHtml) ProcesarVentas(string filePath)
        {
            StringBuilder mensajeTexto = new StringBuilder();
            StringBuilder mensajeHtml = new StringBuilder();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                Dictionary<string, HashSet<string>> erroresSku = new Dictionary<string, HashSet<string>>();
                Dictionary<string, List<string>> otrosErrores = new Dictionary<string, List<string>>();

                for (int row = 2; row <= rowCount; row++)
                {
                    string mensajeErrorColumna = worksheet.Cells[row, 11].Text.Trim();
                    string systemID = worksheet.Cells[row, 1].Text.Trim();
                    string skuCode = worksheet.Cells[row, 2].Text.Trim();
                    string saleDateCell = worksheet.Cells[row, 4].ToString();
                    string Amount = worksheet.Cells[row, 6].Text.Trim();
                    string Descrip1 = worksheet.Cells[row, 7].Text.Trim();
                    string Descrip2 = worksheet.Cells[row, 8].Text.Trim();

                    string mensajeError = ObtenerMensajeError(mensajeErrorColumna, row, systemID, skuCode, saleDateCell, Amount, Descrip1, Descrip2);

                    if (!string.IsNullOrEmpty(mensajeError))
                    {
                        if (mensajeError == "Se detectó un código inválido.")
                        {
                            if (!erroresSku.ContainsKey(mensajeError))
                            {
                                erroresSku[mensajeError] = new HashSet<string>();
                            }
                            erroresSku[mensajeError].Add(skuCode);
                        }
                        else
                        {
                            string mensajeBase = mensajeError.Split(':')[0].Trim();
                            if (!otrosErrores.ContainsKey(mensajeBase))
                            {
                                otrosErrores[mensajeBase] = new List<string>();
                            }

                            if (mensajeError.StartsWith("Detectamos que se ingresan clientes"))
                            {
                                otrosErrores[mensajeBase].Add(systemID);
                            }
                            else if (mensajeError.StartsWith("Detectamos que se ingresan comas"))
                            {
                                // No se agregan filas
                            }
                            else if (mensajeError.StartsWith("Detectamos que se ingresan ventas con el formato de fecha"))
                            {
                                // No se agregan filas
                            }
                            else
                            {
                                otrosErrores[mensajeBase].Add(row.ToString());
                            }
                        }
                    }
                }

                // Construir mensaje de ventas
                mensajeTexto.AppendLine($"Archivo de ventas:");
                mensajeHtml.AppendLine($"<p><span style='text-decoration: underline;'>Archivo de ventas:</span></p>");

                if (erroresSku.Count > 0 || otrosErrores.Count > 0)
                {
                    mensajeTexto.AppendLine("Se detectaron los siguientes errores en el archivo:");
                    mensajeHtml.AppendLine("<p>Se detectaron los siguientes errores en el archivo:</p>");

                    // Errores de SKU
                    if (erroresSku.Count > 0)
                    {
                        mensajeHtml.Append(ConstruirTablaSku(erroresSku.First().Value));
                        mensajeTexto.AppendLine("---");
                        mensajeHtml.AppendLine("<hr>");
                    }

                    // Otros errores
                    bool primerError = true;
                    if (otrosErrores.Count > 0)
                    {
                        foreach (var error in otrosErrores)
                        {
                            if (primerError)
                            {
                                mensajeTexto.AppendLine($"{error.Key}");
                                mensajeHtml.AppendLine($"<p>{error.Key.Replace(":", ".")}</p>");
                                primerError = false;
                            }
                            else
                            {
                                mensajeTexto.AppendLine("---");
                                mensajeHtml.AppendLine("<hr>");
                                mensajeTexto.AppendLine($"También {error.Key.ToLower().Replace("detectamos que", "detectamos que")}");
                                mensajeHtml.AppendLine($"<p>También {error.Key.Replace(":", ".").ToLower().Replace("detectamos que", "detectamos que")}</p>");
                            }

                            if (error.Key.StartsWith("Detectamos que se ingresan clientes"))
                            {
                                mensajeTexto.AppendLine("   Favor de agregar los siguientes clientes:");
                                mensajeTexto.AppendLine(string.Join(Environment.NewLine, error.Value.Distinct()));
                                mensajeHtml.AppendLine("   <p>Favor de agregar los siguientes clientes:</p>");
                                mensajeHtml.AppendLine($"   <p>{string.Join("<br>", error.Value.Distinct())}</p>");
                            }
                            else if (error.Key.StartsWith("Detectamos que se ingresan ventas con el formato de fecha"))
                            {
                                mensajeTexto.AppendLine("   Favor de corregir y reenviar.");
                                mensajeHtml.AppendLine("   <p>Favor de corregir y reenviar.</p>");
                            }
                            else if (error.Key.StartsWith("Detectamos que se ingresan comas"))
                            {
                                mensajeTexto.AppendLine("   Favor de corregir y reenviar.");
                                mensajeHtml.AppendLine("   <p>Favor de corregir y reenviar.</p>");
                            }
                            else
                            {
                                mensajeTexto.AppendLine("   Favor de corregir en las filas:");
                                mensajeTexto.AppendLine($"    {string.Join(", ", error.Value.Distinct())}");
                                mensajeHtml.AppendLine($"    <p>{string.Join(", ", error.Value.Distinct())}</p>");
                            }
                        }
                    }
                }
                else
                {
                    mensajeTexto.AppendLine("No se detectaron errores en el archivo.");
                    mensajeHtml.AppendLine("<p>No se detectaron errores en el archivo.</p>");
                }
            }

            return (mensajeTexto.ToString(), mensajeHtml.ToString());
        }

        private (string mensajeTexto, string mensajeHtml) ProcesarClientes(string filePath)
        {
            StringBuilder mensajeTexto = new StringBuilder();
            StringBuilder mensajeHtml = new StringBuilder();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                Dictionary<string, List<string>> otrosErrores = new Dictionary<string, List<string>>();

                for (int row = 2; row <= rowCount; row++) // Asumimos que la primera fila son encabezados
                {
                    string mensajeErrorColumna = worksheet.Cells[row, 14].Text.Trim(); // Columna "status" (N)
                    string mensajeError = ObtenerMensajeErrorClientes(mensajeErrorColumna, row);

                    if (!string.IsNullOrEmpty(mensajeError))
                    {
                        string mensajeBase = mensajeError.Split(':')[0].Trim();
                        if (!otrosErrores.ContainsKey(mensajeBase))
                        {
                            otrosErrores[mensajeBase] = new List<string>();
                        }

                        if (mensajeError.StartsWith("Detectamos que en la columna \"fiscalId\""))
                        {
                            // No se agregan filas específicas para este error
                        }
                        else if (mensajeError.StartsWith("Detectamos que se ingresan emails erróneos"))
                        {
                            otrosErrores[mensajeBase].Add(row.ToString());
                        }
                        else if (mensajeError.StartsWith("Detectamos que se ingresan códigos de país erróneos"))
                        {
                            // No se agregan filas específicas para este error
                        }
                    }
                }

                // Construir mensaje de clientes
                mensajeTexto.AppendLine($"Archivo de clientes:");
                mensajeHtml.AppendLine($"<p><span style='text-decoration: underline;'>Archivo de clientes:</span></p>");

                if (otrosErrores.Count > 0)
                {
                    mensajeTexto.AppendLine("Se detectaron los siguientes errores en el archivo:");
                    mensajeHtml.AppendLine("<p>Se detectaron los siguientes errores en el archivo:</p>");

                    // Otros errores
                    bool primerError = true;
                    foreach (var error in otrosErrores)
                    {
                        if (primerError)
                        {
                            mensajeTexto.AppendLine($"{error.Key}");
                            mensajeHtml.AppendLine($"<p>{error.Key.Replace(":", ".")}</p>");
                            primerError = false;
                        }
                        else
                        {
                            mensajeTexto.AppendLine("---");
                            mensajeHtml.AppendLine("<hr>");
                            mensajeTexto.AppendLine($"También {error.Key.ToLower().Replace("detectamos que", "detectamos que")}");
                            mensajeHtml.AppendLine($"<p>También {error.Key.Replace(":", ".").ToLower().Replace("detectamos que", "detectamos que")}</p>");
                        }

                        if (error.Key.StartsWith("Detectamos que se ingresan emails erróneos"))
                        {
                            mensajeTexto.AppendLine("   Favor de corregir en las filas:");
                            mensajeTexto.AppendLine($"    {string.Join(", ", error.Value.Distinct())}");
                            mensajeHtml.AppendLine($"    <p>{string.Join(", ", error.Value.Distinct())}</p>");
                        }
                        else
                        {
                            mensajeTexto.AppendLine("   Favor de corregir y reenviar.");
                            mensajeHtml.AppendLine("   <p>Favor de corregir y reenviar.</p>");
                        }
                    }
                }
                else
                {
                    mensajeTexto.AppendLine("No se detectaron errores en el archivo.");
                    mensajeHtml.AppendLine("<p>No se detectaron errores en el archivo.</p>");
                }
            }

            return (mensajeTexto.ToString(), mensajeHtml.ToString());
        }

        private string ObtenerMensajeErrorClientes(string mensajeErrorColumna, int fila)
        {
            if (mensajeErrorColumna.IndexOf("fiscalid", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que en la columna \"fiscalId\" se ingresan números de cuit o cuils los cuales pueden llegar a contener caracteres no válidos tales como guiones, puntos, espacios, etc. Favor de corregir y reenviar.";

            if (mensajeErrorColumna.IndexOf("invalidEmail", StringComparison.OrdinalIgnoreCase) >= 0)
                return $"Detectamos que se ingresan emails erróneos, o con un formato que no es el de una dirección de correo electrónico. Favor de corregir en las filas: {fila}";

            if (mensajeErrorColumna.IndexOf("countryID", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que se ingresan códigos de país erróneos en la columna \"address.countryID\". Favor de corregir y reenviar.";

            return "";
        }

        private string ConstruirTablaSku(HashSet<string> skuCodes)
        {
            string mensajeTexto = "Detectamos que se ingresan códigos no coincidentes con los del archivo de equivalencias.\r\n" +
                                   "Favor de completar el siguiente cuadro:\r\n\r\n";

            string[,] datos = new string[skuCodes.Count + 1, 3];
            datos[0, 0] = "Partner SKU Code";
            datos[0, 1] = "COD-BRID";
            datos[0, 2] = "Name";

            int i = 1;
            foreach (string sku in skuCodes)
            {
                datos[i, 0] = sku;
                datos[i, 1] = "";
                datos[i, 2] = "";
                i++;
            }

            StringBuilder html = new StringBuilder();
            html.AppendLine(mensajeTexto.Replace("\r\n", "<br>"));
            html.AppendLine("<table border='1' style='border-collapse:collapse;'>");
            html.AppendLine("<tr>");
            for (int k = 0; k < datos.GetLength(1); k++)
            {
                html.AppendLine($"<td style='padding:5px;'>{datos[0, k]}</td>");
            }
            html.AppendLine("</tr>");

            for (int j = 1; j < datos.GetLength(0); j++)
            {
                html.AppendLine("<tr>");
                for (int k = 0; k < datos.GetLength(1); k++)
                {
                    html.AppendLine($"<td style='padding:5px;'>{datos[j, k]}</td>");
                }
                html.AppendLine("</tr>");
            }

            html.AppendLine("</table><br>");
            return html.ToString();
        }

        private string ObtenerMensajeErrorStock(string mensajeErrorColumna, int fila, string skuCode, string quantity)
        {
            if (mensajeErrorColumna.IndexOf("Invalid Code", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que se ingresan códigos no coincidentes con los del archivo de equivalencias. Favor de completar el siguiente cuadro:";

            if (mensajeErrorColumna.IndexOf("Date", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que se ingresan fechas inválidas. Recordar que el formato correcto es fecha corta en DMA (Día/Mes/Año).";

            if (mensajeErrorColumna.IndexOf("For input string", StringComparison.OrdinalIgnoreCase) >= 0 ||
                mensajeErrorColumna.IndexOf("Float", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que se ingresan comas y/o otro carácter incompatible en la columna \"Quantity\".";

            return "";
        }

        private string ObtenerMensajeError(string mensajeErrorColumna, int fila, string systemID, string skuCode, string saleDateCell, string Amount, string Descrip1, string Descrip2)
        {
            if (mensajeErrorColumna.IndexOf("Invalid Code", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Se detectó un código inválido.";

            if (mensajeErrorColumna.IndexOf("Invalid Principal", StringComparison.OrdinalIgnoreCase) >= 0)
                return $"Detectamos que se ingresan clientes que no están registrados en el respectivo archivo de clientes: Código de cliente {systemID}";

            if (mensajeErrorColumna.IndexOf("Date", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (!DateTime.TryParseExact(saleDateCell, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out _))
                {
                    return "Detectamos que se ingresan ventas con el formato de fecha inválido. Recordar que el formato correcto es fecha corta en DMA (Día/Mes/Año).";
                }
            }

            if (mensajeErrorColumna.IndexOf("For input string", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Detectamos que se ingresan comas y/o otro carácter incompatible en la columna Amount.";

            if (mensajeErrorColumna.IndexOf("Duplicate Entry", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Se encontró un registro duplicado.";

            if (mensajeErrorColumna.IndexOf("Unauthorized Discount", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Descuento aplicado sin autorización.";

            return "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_DoubleClick(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal; // Restaurar si ya está maximizado
            }
            else
            {
                this.WindowState = FormWindowState.Maximized; // Maximizar si no lo está
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) // Solo reaccionar al clic izquierdo
            {
                isDragging = true;
                startPoint = new Point(e.X, e.Y); // Guardar la posición inicial del clic
            }
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                // Calcular la nueva posición del formulario
                Point newPoint = this.PointToScreen(new Point(e.X, e.Y));
                this.Location = new Point(newPoint.X - startPoint.X, newPoint.Y - startPoint.Y);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }
    }
}
    

    


