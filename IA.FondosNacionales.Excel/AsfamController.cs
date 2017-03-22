using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;
using IA.FondosNacionales.Entity;

namespace IA.FondosNacionales.Excel
{
    public class AsfamController
    {
        public void Procesar(Asfam a)
        {
            var excelAppOut = new ExcelX.Application();
            excelAppOut.Workbooks.Open(@"C:\Fondos Nacionales\Templates\IF_ASFAM.xlsx");
    
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["15", "X"] = a.NroAsignacionesFamiliaresPagadas.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "X"] = a.NroAfiliados.Replace(".", "").Replace(",", "");
            Salida.Cells["22", "V"] = a.NroEmpresas.Replace(".", "").Replace(",", "");

            //var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Asfam\Preliminar\";
            System.IO.FileAttributes attr;
            try
            {
                attr = System.IO.File.GetAttributes(rutaSalida);
            }
            catch (Exception ex)
            {
                System.IO.Directory.CreateDirectory(rutaSalida);
            }
            //_" + fecha + "
            Salida.SaveAs(rutaSalida + "IFAsfam.xlsx");

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppOut);
        }


        public void ProcesarFondo(Asfam c)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\out\" + periodo + @"\Asfam\Preliminar\IFAsfam.xlsx";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Asfam\";

            excelAppOut.Workbooks.Open(rutaEntrada);
            //"Feb-17"
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["14", "M"] = c.AporteFiscalMes.Replace(".", "").Replace(",", "");
            Salida.Cells["15", "M"] = c.Reintego.Replace(".", "").Replace(",", "");

            Salida.Cells["21", "M"] = c.AsFamTrabajadoresActivosMesActual.Replace(".", "").Replace(",", "");
            Salida.Cells["22", "M"] = c.AsFamPensionadosMesActual.Replace(".", "").Replace(",", "");
            Salida.Cells["23", "M"] = c.AsFamTrabajadoresCesantesMesActual.Replace(".", "").Replace(",", "");
            Salida.Cells["24", "M"] = c.AsFamInstitucionesMesActual.Replace(".", "").Replace(",", "");

            Salida.Cells["27", "M"] = c.AsFamTrabajadoresActivosRetroactivo.Replace(".", "").Replace(",", "");
            Salida.Cells["28", "M"] = c.AsFamPensionadosRetroactivo.Replace(".", "").Replace(",", "");
            Salida.Cells["29", "M"] = c.AsFamTrabajadoresCesantesRetroactivo.Replace(".", "").Replace(",", "");
            Salida.Cells["30", "M"] = c.AsFamInstitucionesRetroactivo.Replace(".", "").Replace(",", "");

            Salida.Cells["32", "M"] = c.DocumentosRevalidados.Replace(".", "").Replace(",", "");

            Salida.Cells["38", "M"] = c.DocumentosCaducados.Replace(".", "").Replace(",", "");
            Salida.Cells["39", "M"] = c.DocumentosAnulados.Replace(".", "").Replace(",", "");

            Salida.Cells["47", "O"] = c.DevolucionDocumentosSAFEMCaducados.Replace(".", "").Replace(",", "");
            Salida.Cells["48", "O"] = c.DevolucionDocumentosSAFEMAnulados.Replace(".", "").Replace(",", "");
            Salida.Cells["49", "O"] = c.DocumentosSAFEMRevalidados.Replace(".", "").Replace(",", "");


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
            Salida.SaveAs(rutaSalida + "IFAsfam_" + fecha + ".xlsx");

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }
    }
}
