using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace loadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class Empleado { 
            public string TipoDoc { get; set; }
            public string NroDoc { get; set; }
            public decimal Sueldo { get; set; }
        }

        List<Empleado> empleados = new List<Empleado>();
        List<string> errores = new List<string>();

        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Archivos Excel (*.xlsx)|*.xlsx*";
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    List<Empleado> lista = LeerExcel(ofd.FileName);
                    dgvDatos.DataSource = lista;

                    dgvDatos.Columns["Sueldo"].DefaultCellStyle.Format = "NO";
                    dgvDatos.Columns["Sueldo"].DefaultCellStyle.FormatProvider = new System.Globalization.CultureInfo("es-CO");
                }
            }
        }

        private List<Empleado> LeerExcel(string ruta)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(ruta))) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int filas = worksheet.Dimension.Rows;
                for(int fila = 2; fila <= filas; fila++)
                {
                    string tipo = worksheet.Cells[fila, 1].Text.Trim().ToUpper();
                    if(tipo != "CC" && tipo != "CE")
                    {
                        MessageBox.Show($"Tipo de documento inválido en fila {fila}. Solo se permite CC o CE.");
                        continue;
                    }
                    empleados.Add(new Empleado
                    {
                        TipoDoc = tipo,
                        NroDoc = worksheet.Cells[fila, 2].Text,
                        Sueldo = decimal.Parse(worksheet.Cells[fila, 3].Text)
                    });
                    decimal sueldo = 0;

                    bool esNumeroValido = decimal.TryParse(
                        worksheet.Cells[fila, 3].Text,
                        System.Globalization.NumberStyles.Any,
                        new System.Globalization.CultureInfo("es-CO"),
                        out sueldo
                        );

                    if (!esNumeroValido)
                    {
                        errores.Add($"Fila {fila}: Sueldo inválido");
                        continue;
                    }
                    if(sueldo < 0)
                    {
                        errores.Add($"Fila {fila}: Sueldo inválido");
                        continue;
                    }
                    if (errores.Any())
                    {
                        MessageBox.Show(string.Join("\n", errores));
                    }
                }
            }
            return empleados;
        }
    }
}
