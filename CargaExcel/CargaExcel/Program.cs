﻿using System;
using System.Collections.Generic;
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

                try
                {
                    Excel.Application xlApp = new Excel.Application();

                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(nameFile);

                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    var manufacturers = new List<Manufacturer>();

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
                    Console.WriteLine("Fabricante    " + " \t + \t Descipción" + new string(' ', 21) + "|" + "Activo");
                    foreach (var m in manufacturers)
                    {
                        m.Descripcion = m.Descripcion + new string(' ', 30 - m.Descripcion.Length);
                        Console.WriteLine(String.Format("{0}\t{1}\t{2}", m.Fabricante, m.Descripcion, m.Activo));
                    }

                    Console.WriteLine("Presione una tecla para detenerlo...");
                    Console.ReadKey();
                    xlWorkbook.Close();
                }
                catch (Exception ex)
                {
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
