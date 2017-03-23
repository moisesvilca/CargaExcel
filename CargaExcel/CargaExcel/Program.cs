using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CargaExcel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string nameFile="";
            OpenFileDialog fileSelectPopUp = new OpenFileDialog();
            fileSelectPopUp.Title = "";
            fileSelectPopUp.InitialDirectory = @"c:\";
            fileSelectPopUp.Filter = "All EXCEL FILES (*.xlsx*)|*.xlsx*|All files (*.*)|*.*";
            fileSelectPopUp.FilterIndex = 2;
            fileSelectPopUp.RestoreDirectory = true;
            if (fileSelectPopUp.ShowDialog() == DialogResult.OK)
            {
                nameFile = fileSelectPopUp.FileName;
                Excel.Application xlApp = new Excel.Application();
                try
                {
                    Console.WriteLine("Procesando...");
                    

                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(nameFile);

                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    var manufacturers = new List<Manufacturer>();

                    if (!xlRange.Cells[1, 1].Value2.ToString().ToUpper().Equals("FABRICANTE") &&
                        !xlRange.Cells[1, 2].Value2.ToString().ToUpper().Equals("DESCRIPCION") &&
                        !xlRange.Cells[1, 3].Value2.ToString().ToUpper().Equals("ACTIVO"))
                    {
                        Console.WriteLine("El archivo seleccionado no tiene el formato correcto...");
                    }
                    else
                    {
                        for (int i = 2; i <= xlRange.Rows.Count; i++)
                        {
                            var manufacturer = new Manufacturer();
                            for (int j = 1; j <= xlRange.Columns.Count; j++)
                            {
                                if (xlRange.Cells[i, j] == null || xlRange.Cells[i, j].Value2 == null) continue;
                                switch (j)
                                {
                                    case 1:
                                        manufacturer.Fabricante = xlRange.Cells[i, j].Value2.ToString();
                                        break;
                                    case 2:
                                        manufacturer.Descripcion = xlRange.Cells[i, j].Value2.ToString();
                                        break;
                                    case 3:
                                        manufacturer.Activo = xlRange.Cells[i, j].Value2.ToString();
                                        break;
                                }
                            }
                            if (!string.IsNullOrEmpty(manufacturer.Fabricante))
                                manufacturers.Add(manufacturer);
                        }
                        Console.Clear();
                        Console.WriteLine("Fabricante    " + " \t  \t Descipción" + new string(' ', 10) + "Activo");
                        foreach (var m in manufacturers)
                        {
                            m.Descripcion = m.Descripcion + new string(' ', 30 - m.Descripcion.Length);
                            Console.WriteLine(String.Format("{0}\t\t{1}\t{2}", m.Fabricante, m.Descripcion, m.Activo));
                        }


                    }
                    Console.WriteLine("Presione una tecla para detenerlo...");
                    Console.ReadKey();
                    xlWorkbook.Close();
                }
                catch (Exception ex)
                {
                    xlApp.Workbooks.Close();
                    throw;
                }
            }
            else
            {
                Console.WriteLine("Debe seleccionar un archivo excel...");
                Console.ReadKey();
            }
        }

        public class Manufacturer
        {
            public string Fabricante { get; set; }
            public string Descripcion { get; set; }
            public string Activo { get; set; }

        }

    }
}
