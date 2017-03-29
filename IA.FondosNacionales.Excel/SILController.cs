using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;
using IA.FondosNacionales.Entity;

namespace IA.FondosNacionales.Excel
{
    public class SILController
    {
        public void Procesar(SIL s, string periodo)
        {
            var excelAppOut = new ExcelX.Application();

            var rutaEntrada = @"C:\Fondos Nacionales\Templates\IF_SIL";
            
            Utilidades.AbrirLibro(excelAppOut, rutaEntrada);
           
            

            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["12", "U"] = s.NumSubsidiosIniciados.Replace(".", "").Replace(",", "");
            Salida.Cells["13", "U"] = s.NumAfiliadosCotizantes.Replace(".", "").Replace(",", "");

            Salida.Cells["14", "U"] = s.NumEmpresasCotizantes.Replace(".", "").Replace(",", "");
            Salida.Cells["15", "U"] = s.NumTrabajadoresAfiliados.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "U"] = s.NumEmpresasAfiliadas.Replace(".", "").Replace(",", "");

            Salida.Cells["52", "Q"] = s.ValorNotaInterna.Replace(".", "").Replace(",", "");

            //var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Sil\Preliminar\";
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
            Salida.SaveAs(rutaSalida + "IFSIL" + Utilidades.ExtensionLibro(Salida.Application.ActiveWorkbook));
            
            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);
        }

        public void ProcesarFondo(SIL c, string periodo)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\out\" + periodo + @"\Sil\Preliminar\IFSIL";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Sil\";

            Utilidades.AbrirLibro(excelAppOut, rutaEntrada);

         
            
            //"Feb-17"
            ExcelX._Worksheet Salida = (ExcelX.Worksheet)excelAppOut.Sheets["Template"];
            Salida.Cells["12", "I"] = c.Cotizaciones.Replace(".", "").Replace(",", "");
            Salida.Cells["13", "I"] = c.CotizacionesPeriodosAnteriores.Replace(".", "").Replace(",", "");
            Salida.Cells["14", "I"] = c.ReajusteLey17332.Replace(".", "").Replace(",", "");
            Salida.Cells["15", "I"] = c.CotizacionesEntidadesPagadorasdeSubsidios.Replace(".", "").Replace(",", "");
            Salida.Cells["16", "I"] = c.ReintegroporCobroIndebidodeSubsidio.Replace(".", "").Replace(",", "");


            Salida.Cells["23", "I"] = c.SILEnfermedadOrigenComun.Replace(".", "").Replace(",", "");
            Salida.Cells["24", "I"] = c.SILSubsidioMaternalSuplementario.Replace(".", "").Replace(",", "");

            Salida.Cells["26", "K"] = c.DescuentoBeneficiosNoCobrados.Replace(".", "").Replace(",", "");

            Salida.Cells["29", "I"] = c.SREnfermedadOrigenComun.Replace(".", "").Replace(",", "");
            Salida.Cells["30", "I"] = c.SRSubsidioMaternalSuplementario.Replace(".", "").Replace(",", "");

            Salida.Cells["33", "I"] = c.CFPEnfermedadOrigenComun.Replace(".", "").Replace(",", "");
            Salida.Cells["34", "I"] = c.CFPSubsidioMaternalSuplementario.Replace(".", "").Replace(",", "");

            Salida.Cells["37", "I"] = c.CFSEnfermedadOrigenComun.Replace(".", "").Replace(",", "");
            Salida.Cells["38", "I"] = c.CFSSubsidioMaternalSuplementario.Replace(".", "").Replace(",", "");

            Salida.Cells["41", "I"] = c.OCEnfermedadOrigenComun.Replace(".", "").Replace(",", "");
            Salida.Cells["42", "I"] = c.OCSubsidioMaternalSuplementario.Replace(".", "").Replace(",", "");

            Salida.Cells["46", "I"] = c.OtrosEgresos.Replace(".", "").Replace(",", "");
            
            
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
            Salida.SaveAs(rutaSalida + "IFSIL_" + fecha + Utilidades.ExtensionLibro(Salida.Application.ActiveWorkbook));

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }

        

        public void GenerarAnexo(string periodo)
        {
            var excelApp = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\in\" + periodo + @"\SISILHIA";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Sil\Anexo\";
            var rutaTemplate = @"C:\Fondos Nacionales\Templates\ANEXO_SIL";
            var rutaDos = @"C:\Fondos Nacionales\in\" + periodo + @"\ESTEMP";
            var rutaAux = @"C:\Fondos Nacionales\Auxiliar\AUX_DINAMICA";
            
            ExcelX.Workbook libroEntrada = Utilidades.AbrirLibro(excelApp, rutaEntrada);
            ExcelX.Workbook libroDestino = Utilidades.AbrirLibro(excelApp, rutaTemplate);
            

            //Primero
            ExcelX._Worksheet Cuadro1 = libroEntrada.Sheets["CUADRO N° 1"];
            ExcelX._Worksheet Cuadro1_a6 = libroDestino.Sheets["SIL Anexo 6 - Cuadro Nº1"];

            var from = Cuadro1.Range["D12:F13"];
            var to = Cuadro1_a6.Range["E22:G23"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["H12:J13"];
            to = Cuadro1_a6.Range["I22:K23"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["D16:F17"];
            to = Cuadro1_a6.Range["E25:G26"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["H16:J17"];
            to = Cuadro1_a6.Range["I25:K26"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["D20:F21"];
            to = Cuadro1_a6.Range["E28:G29"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["H20:J21"];
            to = Cuadro1_a6.Range["I28:K29"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            
            //Segundo
            ExcelX._Worksheet Cuadro2a = libroEntrada.Sheets["CUADRO N°2-A"];
            ExcelX._Worksheet Cuadro2b = libroEntrada.Sheets["CUADRO N°2-B"];
            ExcelX._Worksheet Cuadro2_a6 = libroDestino.Sheets["SIL Anexo 6 - Cuadro N° 2-A y B"];

            from = Cuadro2a.Range["C10:J37"];
            to = Cuadro2_a6.Range["D18:K45"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro2b.Range["C10:J37"];
            to = Cuadro2_a6.Range["D59:K86"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //Tercero
            ExcelX._Worksheet Cuadro3 = libroEntrada.Sheets["CUADRO N°3"];
            ExcelX._Worksheet Cuadro3_a6 = libroDestino.Sheets["SIL Anexo 6 - Cuadro Nº 3"];

            from = Cuadro3.Range["C10:J24"];
            to = Cuadro3_a6.Range["D18:K32"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            //cuarto
            ExcelX._Worksheet Cuadro4 = libroEntrada.Sheets["CUADRO N° 4"];
            ExcelX._Worksheet Cuadro4_a6 = libroDestino.Sheets["SIL Anexo 6 - Cuadro Nº 4"];

            from = Cuadro4.Range["C10:J19"];
            to = Cuadro4_a6.Range["D21:K30"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //Quinto
            ExcelX._Worksheet Cuadro5 = libroEntrada.Sheets["CUADRO N°5"];
            ExcelX._Worksheet Cuadro5_a6 = libroDestino.Sheets["SIL Anexo 6 - Cuadro Nº5"];

            from = Cuadro5.Range["D15:E29"];
            to = Cuadro5_a6.Range["D19:E33"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            ///////////////////////////////
            ExcelX.Workbook libroDos = Utilidades.AbrirLibro(excelApp, rutaDos);
            ExcelX.Workbook libroAux = Utilidades.AbrirLibro(excelApp, rutaAux);
            //Quinto
            ExcelX._Worksheet Cuadro6 = libroDos.Sheets["Cuadro 6"];
            ExcelX._Worksheet Cuadro6y7_a6 = libroDestino.Sheets["SIL Anexo 6-Cuadro Nº 6 Y 7"];

            from = Cuadro6.Range["E12:E26"];
            to = Cuadro6y7_a6.Range["D16:D30"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //Quinto
            ExcelX._Worksheet Cuadro7 = libroDos.Sheets["Cuadro 7"];
            ExcelX._Worksheet auxiliar = libroAux.Sheets["aux"];

            from = Cuadro7.Range["M12:M27"];
            to = auxiliar.Range["M12:M27"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro7.Range["P12:P27"];
            to = auxiliar.Range["P12:P27"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            ExcelX.PivotTables pts = (ExcelX.PivotTables)auxiliar.PivotTables(Type.Missing);

            var ctn = pts.Count;
            pts.Item(1).RefreshTable();
            
            from = auxiliar.Range["S13:T22"];
            to = Cuadro6y7_a6.Range["D47:E56"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //Resumen
            ExcelX._Worksheet Resumen = libroEntrada.Sheets["ANEXO N° 3"];
            ExcelX._Worksheet Resumen_a6 = libroDestino.Sheets["Resumen Cotizaciones"];

            from = Resumen.Range["C12:E19"];
            to = Resumen_a6.Range["D16:F23"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Resumen.Range["C25:E28"];
            to = Resumen_a6.Range["D28:F31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Resumen.Range["C34:E35"];
            to = Resumen_a6.Range["D35:F36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Resumen.Range["C44:E45"];
            to = Resumen_a6.Range["D42:F43"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Resumen.Range["C46:E46"];
            to = Resumen_a6.Range["D45:F45"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Resumen.Range["C47:E47"];
            to = Resumen_a6.Range["D44:F44"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            
            from = Resumen.Range["C48:E49"];
            to = Resumen_a6.Range["D46:F47"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //cerrando
            System.IO.FileAttributes attr;
            try
            {
                attr = System.IO.File.GetAttributes(rutaSalida);
            }
            catch (Exception ex)
            {
                System.IO.Directory.CreateDirectory(rutaSalida);
            }


            libroDestino.SaveAs(rutaSalida + "Anexo_SIL_" + fecha + Utilidades.ExtensionLibro(libroDestino));

            libroDestino.Close(false);
            libroEntrada.Close(false);
            libroAux.Close(false);
            libroDos.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        }

    }
}
