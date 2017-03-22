using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;
using IA.FondosNacionales.Entity;

namespace IA.FondosNacionales.Excel
{
    public class MaternalController
    {
        public void ProcesarFondo(Maternal m)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\Templates\IF_MATERNAL.xlsx";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Maternal\";

            excelAppOut.Workbooks.Open(rutaEntrada);
            //"Feb-17"
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];

            Salida.Cells["14", "K"] = m.A1.Replace(".", "").Replace(",", "");
            Salida.Cells["15", "K"] = m.A2.Replace(".", "").Replace(",", "");
            Salida.Cells["18", "I"] = m.A31.Replace(".", "").Replace(",", "");
            Salida.Cells["19", "I"] = m.A32.Replace(".", "").Replace(",", "");
            Salida.Cells["22", "I"] = m.A41.Replace(".", "").Replace(",", "");
            Salida.Cells["23", "I"] = m.A42.Replace(".", "").Replace(",", "");

            Salida.Cells["30", "K"] = m.C1.Replace(".", "").Replace(",", "");
            Salida.Cells["31", "K"] = m.C2.Replace(".", "").Replace(",", "");
            Salida.Cells["32", "K"] = m.C3.Replace(".", "").Replace(",", "");
            Salida.Cells["33", "K"] = m.C4.Replace(".", "").Replace(",", "");
            Salida.Cells["34", "K"] = m.C5.Replace(".", "").Replace(",", "");

            Salida.Cells["37", "I"] = m.C61.Replace(".", "").Replace(",", "");
            Salida.Cells["38", "I"] = m.C62.Replace(".", "").Replace(",", "");
            Salida.Cells["39", "I"] = m.C63.Replace(".", "").Replace(",", "");
            Salida.Cells["40", "I"] = m.C64.Replace(".", "").Replace(",", "");
            Salida.Cells["41", "I"] = m.C65.Replace(".", "").Replace(",", "");

            Salida.Cells["44", "I"] = m.C71.Replace(".", "").Replace(",", "");
            Salida.Cells["45", "I"] = m.C72.Replace(".", "").Replace(",", "");
            Salida.Cells["46", "I"] = m.C73.Replace(".", "").Replace(",", "");
            Salida.Cells["47", "I"] = m.C74.Replace(".", "").Replace(",", "");
            Salida.Cells["48", "I"] = m.C75.Replace(".", "").Replace(",", "");

            Salida.Cells["51", "I"] = m.C81.Replace(".", "").Replace(",", "");
            Salida.Cells["52", "I"] = m.C82.Replace(".", "").Replace(",", "");
            Salida.Cells["53", "I"] = m.C83.Replace(".", "").Replace(",", "");
            Salida.Cells["54", "I"] = m.C84.Replace(".", "").Replace(",", "");
            Salida.Cells["55", "I"] = m.C85.Replace(".", "").Replace(",", "");

            Salida.Cells["58", "I"] = m.C91.Replace(".", "").Replace(",", "");
            Salida.Cells["59", "I"] = m.C92.Replace(".", "").Replace(",", "");
            Salida.Cells["60", "I"] = m.C93.Replace(".", "").Replace(",", "");
            Salida.Cells["61", "I"] = m.C94.Replace(".", "").Replace(",", "");
            Salida.Cells["62", "I"] = m.C95.Replace(".", "").Replace(",", "");

            Salida.Cells["68", "K"] = m.E1.Replace(".", "").Replace(",", "");
            Salida.Cells["69", "K"] = m.E2.Replace(".", "").Replace(",", "");
            Salida.Cells["70", "K"] = m.E3.Replace(".", "").Replace(",", "");
            Salida.Cells["71", "K"] = m.E4.Replace(".", "").Replace(",", "");
            Salida.Cells["72", "K"] = m.E5.Replace(".", "").Replace(",", "");


            Salida.Cells["76", "K"] = m.F1.Replace(".", "").Replace(",", "");
            Salida.Cells["77", "K"] = m.F2.Replace(".", "").Replace(",", "");
            Salida.Cells["78", "K"] = m.F3.Replace(".", "").Replace(",", "");
            Salida.Cells["79", "K"] = m.F4.Replace(".", "").Replace(",", "");
            Salida.Cells["80", "K"] = m.F5.Replace(".", "").Replace(",", "");

            Salida.Cells["84", "K"] = m.G1.Replace(".", "").Replace(",", "");
            Salida.Cells["85", "K"] = m.G2.Replace(".", "").Replace(",", "");
            Salida.Cells["86", "K"] = m.G3.Replace(".", "").Replace(",", "");
            Salida.Cells["87", "K"] = m.G4.Replace(".", "").Replace(",", "");



            System.IO.FileAttributes attr;
            try
            {
                attr = System.IO.File.GetAttributes(rutaSalida);
            }
            catch (Exception ex)
            {
                System.IO.Directory.CreateDirectory(rutaSalida);
            }

            Salida.Name = periodo;
            Salida.SaveAs(rutaSalida + "IFMaternal_" + fecha + ".xlsx");

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }
    }
}
