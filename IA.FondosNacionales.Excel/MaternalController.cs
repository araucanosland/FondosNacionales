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
        public void ProcesarFondo(Maternal m, string periodo)
        {
            var excelAppOut = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\Templates\IF_MATERNAL";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Maternal\";


            Utilidades.AbrirLibro(excelAppOut, rutaEntrada);
            
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
            Salida.SaveAs(rutaSalida + "IFMaternal_" + fecha + Utilidades.ExtensionLibro(Salida.Application.ActiveWorkbook));

            excelAppOut.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppOut);


        }


        public void GenerarAnexo(string periodo)
        {
            var excelApp = new ExcelX.Application();
            var fecha = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "");
            //var periodo = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0');
            var rutaEntrada = @"C:\Fondos Nacionales\in\" + periodo + @"\ANEXO4";
            var rutaSalida = @"C:\Fondos Nacionales\out\" + periodo + @"\Maternal\Anexo\";
            var rutaTemplate = @"C:\Fondos Nacionales\Templates\ANEXO_MATERNAL";



            ExcelX.Workbook libroEntrada = Utilidades.AbrirLibro(excelApp, rutaEntrada);
            ExcelX.Workbook libroDestino = Utilidades.AbrirLibro(excelApp, rutaTemplate);
            
            //Primero
            ExcelX._Worksheet Cuadro1 = libroEntrada.Sheets["ResCotizPrevi"];
            ExcelX._Worksheet Cuadro1_a6 = libroDestino.Sheets["Anexo 4-Res"];

            var from = Cuadro1.Range["D11:D20"];
            var to = Cuadro1_a6.Range["D13:D22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["F11:F20"];
            to = Cuadro1_a6.Range["E13:E22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H11:H20"];
            to = Cuadro1_a6.Range["F13:F22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["I11:I20"];
            to = Cuadro1_a6.Range["G13:G22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["K11:K20"];
            to = Cuadro1_a6.Range["H13:H22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["L11:L20"];
            to = Cuadro1_a6.Range["I13:I22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["M11:M20"];
            to = Cuadro1_a6.Range["J13:J22"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //SPRT
            from = Cuadro1.Range["D24:D29"];
            to = Cuadro1_a6.Range["D26:D31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //TO DO: Mirar por que no copia estos valores, aparte no suma la caja de total
            from = Cuadro1.Range["F24:F29"];
            to = Cuadro1_a6.Range["E26:E31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H24:H29"];
            to = Cuadro1_a6.Range["F26:F31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["I24:I29"];
            to = Cuadro1_a6.Range["G26:G31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["K24:K29"];
            to = Cuadro1_a6.Range["H26:H31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["L24:L29"];
            to = Cuadro1_a6.Range["I26:I31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["M24:M29"];
            to = Cuadro1_a6.Range["J26:J31"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
         
            //TO DO: Revisar por que no copia
            from = Cuadro1.Range["D33:D34"];
            to = Cuadro1_a6.Range["D35:D36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["F33:F34"];
            to = Cuadro1_a6.Range["E35:E36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H33:H34"];
            to = Cuadro1_a6.Range["F35:F36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["I33:I34"];
            to = Cuadro1_a6.Range["G35:G36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["K33:K34"];
            to = Cuadro1_a6.Range["H35:H36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //SPRT
            from = Cuadro1.Range["D41:D46"];
            to = Cuadro1_a6.Range["D35:D36"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["F41:F46"];
            to = Cuadro1_a6.Range["E42:E27"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H41:H46"];
            to = Cuadro1_a6.Range["F42:F47"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            //from = Cuadro1.Range["I41:I46"];
            //to = Cuadro1_a6.Range["G42:G37"];
            //from.Copy();
            //to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //from = Cuadro1.Range["K41:K46"];
            //to = Cuadro1_a6.Range["H42:H27"];
            //from.Copy();
            //to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //SPRT
            from = Cuadro1.Range["D50"];
            to = Cuadro1_a6.Range["D51"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["F50"];
            to = Cuadro1_a6.Range["E51"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H50"];
            to = Cuadro1_a6.Range["F51"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            //from = Cuadro1.Range["I41:I46"];
            //to = Cuadro1_a6.Range["G42:G37"];
            //from.Copy();
            //to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //from = Cuadro1.Range["K41:K46"];
            //to = Cuadro1_a6.Range["H42:H27"];
            //from.Copy();
            //to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            //SPRT
            from = Cuadro1.Range["D54:D60"];
            to = Cuadro1_a6.Range["D55:D61"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["F54:F60"];
            to = Cuadro1_a6.Range["E55:E61"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["H54:H60"];
            to = Cuadro1_a6.Range["F55:F61"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            from = Cuadro1.Range["I54:I60"];
            to = Cuadro1_a6.Range["G55:G61"];
            from.Copy();
            to.PasteSpecial(ExcelX.XlPasteType.xlPasteValues, ExcelX.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);


            from = Cuadro1.Range["K54:K60"];
            to = Cuadro1_a6.Range["H55:H61"];
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


            libroDestino.SaveAs(rutaSalida + "Anexo_Maternal_" + fecha + Utilidades.ExtensionLibro(libroDestino));

            libroDestino.Close(false);
            libroEntrada.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        }

    }
}
