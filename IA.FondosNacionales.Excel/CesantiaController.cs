using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;
using IA.FondosNacionales.Entity;

namespace IA.FondosNacionales.Excel
{
    public class CesantiaController
    {
        public void Procesar(Cesantia c)
        {
            var excelAppOut = new ExcelX.Application();
            //var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":","").Replace(" ","");
            var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Cesantia\Preliminar\";
            

            excelAppOut.Workbooks.Open(@"C:\Fondos Nacionales\Templates\IF_CESANTIA.xls");
            //"Feb-17"
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["14", "V"] = c.NroSubsidios.Replace(".","").Replace(",", "");
            Salida.Cells["15", "V"] = c.NroAfiliados.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "V"] = "133";
            Salida.Cells["21", "T"] = c.NroEmpresas.Replace(".", "").Replace(",", "");

            
            System.IO.FileAttributes attr;
            try
            {
                attr = System.IO.File.GetAttributes(rutaSalida);
            }
            catch(Exception ex)
            {
                System.IO.Directory.CreateDirectory(rutaSalida);    
            }
            //_" + fecha + "
            Salida.SaveAs(rutaSalida + "IFCesantia.xls");
            
            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);
        }


        public void ProcesarFondo(Cesantia c)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":","").Replace(" ","");
            var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\out\" + periodo + @"\Cesantia\Preliminar\IFCesantia.xls";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Cesantia\";

            excelAppOut.Workbooks.Open(rutaEntrada);
            //"Feb-17"
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["14", "I"] = c.AporteFiscalMes.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "I"] = c.Reintego.Replace(".", "").Replace(",", "");
            Salida.Cells["27", "I"] = c.SubsidiosCesantia.Replace(".", "").Replace(",", "");

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
            Salida.SaveAs(rutaSalida + "IFCesantia_" + fecha + ".xls");

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }
    }
}
