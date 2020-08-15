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

namespace HowToReadExcelWithEpplus
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {

            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                textBox1.Text = openFileDialog1.FileName;

                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                var package = new ExcelPackage(new FileInfo(openFileDialog1.FileName));

                int contadorMes = 1;
                int contadorHoja = 1;
                int contadorServiciosPorMes = 0;
                int litrosPorMes = 0;



                List<ServicioDTO> servicios = new List<ServicioDTO>();
                foreach (var item in package.Workbook.Worksheets)
                {

                    int ColumnaInicio = 1;
                    int ColumnaFin = 5;
                    int ColumnaConsumo = 7;
                    int ColumnaVale = 10;
                    int maximoRegistros = 100;





                    string valeBase = "";
                    string valeCero = "";
                    //first the males. start at row 2 to skip the header
                    for (int row = 1; row <= maximoRegistros; row++)
                    {


                        if (row == 1)
                        {
                            string textoConsumo = item.Cells[row, ColumnaConsumo].Value.ToString();
                            if (textoConsumo.ToUpper().Trim() == "CONSUMO")
                            {
                                //MessageBox.Show("Encontrada columna consumo");
                            }

                            string textoVale = item.Cells[row, ColumnaVale].Value.ToString();
                            if (textoVale.ToUpper().Trim() == "NUMERO DE VALE")
                            {
                                //MessageBox.Show("Encontrada columna vale");
                            }
                            else
                            {
                                ColumnaVale = 11;

                                textoVale = item.Cells[row, ColumnaVale].Value.ToString();
                                if (textoVale.ToUpper().Trim() == "NUMERO DE VALE")
                                {
                                    //MessageBox.Show("Encontrada columna vale");
                                }
                            }
                        }


                        if (row > 1)
                        {
                            if (item.Cells[row, ColumnaVale].Value != null)
                            {
                                var tmpVale = item.Cells[row, ColumnaVale].Value.ToString();
                                if (valeBase != tmpVale)
                                {
                                    valeBase = item.Cells[row, ColumnaVale].Value.ToString();
                                    contadorServiciosPorMes++;

                                }

                                if (item.Cells[row, ColumnaConsumo+1].Value != null)
                                {
                                    if (item.Cells[row, ColumnaConsumo].Value != null)
                                    {
                                        var tmpLitros = item.Cells[row, ColumnaConsumo].Value.ToString();
                                        bool res;
                                        int a;
                                        string myStr = "12";
                                        res = int.TryParse(tmpLitros, out a);

                                        if (res)
                                        {
                                            var tmpNumeroLirtos = int.Parse(tmpLitros);
                                            litrosPorMes = litrosPorMes + tmpNumeroLirtos;
                                        }

                                    }
                                }
                               

                                //else
                                //{
                                //    var tmpLitros = item.Cells[row, ColumnaConsumo].Value.ToString();
                                //    var tmpNumeroLirtos = int.Parse(tmpLitros);
                                //    litrosPorMes = litrosPorMes + tmpNumeroLirtos;
                                //}


                                if (tmpVale == string.Empty)
                                {
                                    //break;
                                }
                            }






                        }


                    }

                    Cursor.Current = Cursors.Default;


                    if (contadorHoja % 3 == 0)
                    {

                        ServicioDTO conteoDTO = new ServicioDTO();
                        conteoDTO.Anio = 2019;
                        if (contadorHoja == 13)
                        {
                            conteoDTO.Anio = 2020;
                        }
                        conteoDTO.Mes = contadorMes;
                        conteoDTO.Numero = contadorMes; ;
                        conteoDTO.Servicios = contadorServiciosPorMes;
                        conteoDTO.Litros = litrosPorMes;
                        contadorServiciosPorMes = 0;
                        litrosPorMes = 0;


                        servicios.Add(conteoDTO);
                        contadorMes++;



                    }
                    contadorHoja++;


                }


                Cursor.Current = Cursors.Default;
                dataGridView1.DataSource = servicios;
                dataGridView1.Refresh();



            }

        }



    }
}
