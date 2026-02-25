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

                    dgvDatos.Columns["Sueldo"].DefaultCellStyle.Format = "No";
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
                    decimal totalNomina = empleados.Sum(e => e.Sueldo);
                    MessageBox.Show(totalNomina.ToString("C0", new System.Globalization.CultureInfo("es-CO")));
                    var agrupado = empleados
                        .GroupBy(e => e.TipoDoc)
                        .Select(g => new
                        {
                            TipoDoc = g.Key,
                            cantidad = g.Count(),
                            Total = g.Sum(x => x.Sueldo)
                        })
                        .ToList();
                    dgvResumen.DataSource = agrupado;
                    var estadisticas = new
                    {
                        Promedio = empleados.Average(e => e.Sueldo),
                        Maximo = empleados.Max(e => e.Sueldo),
                        Minimo = empleados.Min(e => e.Sueldo),
                        Total = empleados.Sum(e => e.Sueldo),
                        Cantidad = empleados.Count()
                    };
                    dgvEstadisticas.DataSource = new List<object> { estadisticas };
                    var duplicados = empleados
                        .GroupBy(e => e.NroDoc)
                        .Where(g =>  g.Count() > 1)
                        .Select(g => g.Key)
                        .ToList();
                    if (duplicados.Any())
                    {
                        MessageBox.Show("Documentos duplicados:\n" + string.Join("\n", duplicados));
                    }
                    else
                    {
                        MessageBox.Show("No hay duplicados");
                    }
                }
            }
            return empleados;
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            ExportarCSV(empleados);
        }

        private void ExportarCSV(List<Empleado> empleados)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV (*.csv)|*.csv*";
                sfd.FileName = "Nomina.csv";
                if (sfd.ShowDialog() == DialogResult.OK) { 
                    var lineas = new List<string>();
                    lineas.Add("tipo_doc,nro_doc,sueldo");
                    foreach (var emp in empleados)
                    {
                        lineas.Add($"{emp.TipoDoc},{emp.NroDoc}, {emp.Sueldo}");
                    }
                    File.WriteAllLines(sfd.FileName, lineas);
                    MessageBox.Show("Archivo CSV exportado correctamente");
                }
            }
        }
    }
}
