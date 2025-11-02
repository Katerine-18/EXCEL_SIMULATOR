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

namespace EXCEL_SIMULATOR
{
    public partial class FormVentana : Form
    {
        
        public FormVentana()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void btnMinizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void FormVentana_Load(object sender, EventArgs e)
        {
            dgvCeldas.Columns.Clear();

            //Definiendo variables para las columnas y filas
            int numberOfColumns = 26; // Columnas de la A a la Z
            for (int i = 0; i < numberOfColumns; i++)
            {
                //Nombre de las columnas
                char columnNameChar = (char)('A' + i);
                string columnName = columnNameChar.ToString();

                //Agregar columnas al DataGridView
                if (i >= 26)
                {
                    int primerCaracter = (i / 26) - 1;
                    char primerCaracterColumna = (char)('A' + primerCaracter);
                    int segundoCaracter = i % 26;
                    char segundoCaracterColumna = (char)('A' + segundoCaracter);
                    columnName = primerCaracterColumna.ToString() + segundoCaracterColumna.ToString();
                }

                dgvCeldas.Columns.Add(columnName, columnName);
            }

            // Agregar 50 filas
            int numberOfRows = 50;
            for (int i = 0; i < numberOfRows; i++)
            {
                dgvCeldas.Rows.Add();
            }
        }


        private void dgvCeldas_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            int rowNumero = e.RowIndex + 1;

            // Obtener el área del encabezado de fila
            Rectangle rowHeaderBounds = new Rectangle(
                e.RowBounds.Left,
                e.RowBounds.Top,
                ((DataGridView)sender).RowHeadersWidth,
                e.RowBounds.Height);

            // Centrar el texto vertical y horizontalmente
            StringFormat format = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            e.Graphics.DrawString(
                rowNumero.ToString(),
                this.Font,
                SystemBrushes.ControlText,
                rowHeaderBounds,
                format);
        }

        private void dgvCeldas_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Doble click al seleccionar celda
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgvCeldas.BeginEdit(true);
            }
        }

        private Dictionary<string, string> formulas = new Dictionary<string, string>();

        private void dgvCeldas_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string cellKey = $"{dgvCeldas.Columns[e.ColumnIndex].Name}{e.RowIndex + 1}";
                if (formulas.ContainsKey(cellKey))
                {
                    txtFormulaBar.Text = formulas[cellKey];
                }
                else
                {
                   var cellValue = dgvCeldas.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    txtFormulaBar.Text=cellValue?.ToString() ?? string.Empty;

                }

            }
        }

        private void dgvCeldas_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var cell = dgvCeldas.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string cellkey = $"{dgvCeldas.Columns[e.ColumnIndex].Name}{e.RowIndex + 1}";
                string input = cell.Value?.ToString() ?? string.Empty;

                //Si empieza con "=" es una formula

                if (input.StartsWith("="))
                {
                    formulas[cellkey] = input;

                    //Mostrar el resultado de la formula
                    double resultado = EvaluarFormula(input);
                    cell.Value = resultado;

                    //Actualizar la barra de formulas
                    txtFormulaBar.Text = input;
                }
                else
                {
                    if (formulas.ContainsKey(cellkey))
                        formulas.Remove(cellkey);

                }



            }


        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Guardar archivo Excel";
            saveFileDialog.FileName = "MiDocumento.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    GuardarExcel(saveFileDialog.FileName);
                    MessageBox.Show("Archivo guardado exitosamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al guardar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }

        }


        //Creando metodos para evaluar formulas simples

        private double EvaluarFormula(string formula)
        {
            try
            {
                //quitar el "=" al inicio
                formula = formula.Substring(1);

                //Reemplazar referencias de celdas por sus valores
                foreach (DataGridViewColumn col in dgvCeldas.Columns) // Corregido el nombre de la clase aquí
                {
                    for (int row = 0; row < dgvCeldas.Rows.Count - 1; row++)
                    {
                        string cellRef = $"{col.Name}{row + 1}";
                        if (formula.Contains(cellRef))
                        {
                            var cellValue = dgvCeldas.Rows[row].Cells[col.Index].Value;
                            double valor = 0;
                            if (cellValue != null && double.TryParse(cellValue.ToString(), out valor))
                            {
                                formula = formula.Replace(cellRef, valor.ToString());
                            }
                            else
                            {
                                formula = formula.Replace(cellRef, "0");
                            }
                        }
                    }
                }

                var resultado = new DataTable().Compute(formula, null);
                return Convert.ToDouble(resultado);
            }
            catch
            {
                return 0;
            }
        }

        private void GuardarExcel(string rutaArchivo)
        {
            

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Hoja1");

                //Guardando encabezados de columna
                for (int col = 0; col < dgvCeldas.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dgvCeldas.Columns[col].HeaderText;

                }

                //Guardando datos en las celdas
                for (int row = 0; row < dgvCeldas.Rows.Count; row++)
                {
                    for (int col = 0; col < dgvCeldas.Columns.Count; col++)
                    {
                        var cellValue = dgvCeldas.Rows[row].Cells[col].Value;
                        if (cellValue != null)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = cellValue.ToString();

                        }
                    }
                }

                worksheet.Cells.AutoFitColumns();

                //Archivo Guardado
                FileInfo fileInfo = new FileInfo(rutaArchivo);
                package.SaveAs(fileInfo);

            }
        }

    }
} 


    
