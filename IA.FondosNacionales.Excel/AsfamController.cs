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
        public void Procesar(Asfam a, string periodo)
        {
            var excelAppOut = new ExcelX.Application();

            Utilidades.AbrirLibro(excelAppOut, @"C:\Fondos Nacionales\Templates\IF_ASFAM");
    
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["15", "X"] = a.NroAsignacionesFamiliaresPagadas.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "X"] = a.NroAfiliados.Replace(".", "").Replace(",", "");
            Salida.Cells["22", "V"] = a.NroEmpresas.Replace(".", "").Replace(",", "");

            Salida.Cells["57", "U"] = a.NI_Tramo0.Replace(".", "").Replace(",", "");
            Salida.Cells["58", "U"] = a.NI_Tramo1.Replace(".", "").Replace(",", "");
            Salida.Cells["59", "U"] = a.NI_Tramo2.Replace(".", "").Replace(",", "");
            Salida.Cells["60", "U"] = a.NI_Tramo3.Replace(".", "").Replace(",", "");
            Salida.Cells["61", "U"] = a.NI_Tramo4.Replace(".", "").Replace(",", "");


            //var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
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
            Salida.SaveAs(rutaSalida + "IFAsfam" + Utilidades.ExtensionLibro(Salida.Application.ActiveWorkbook));

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppOut);
        }


        public void ProcesarFondo(Asfam c, string periodo)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\out\" + periodo + @"\Asfam\Preliminar\IFAsfam";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Asfam\";

            Utilidades.AbrirLibro(excelAppOut, rutaEntrada);
            
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
            Salida.SaveAs(rutaSalida + "IFAsfam_" + fecha + Utilidades.ExtensionLibro(Salida.Application.ActiveWorkbook));

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }
    }
}
