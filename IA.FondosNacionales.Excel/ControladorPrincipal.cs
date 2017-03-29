using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;
using IA.FondosNacionales.Entity;

namespace IA.FondosNacionales.Excel
{
    public class ControladorPrincipal
    {

        private CesantiaController cc;
        private AsfamController ac;
        private SILController sc;
        private MaternalController mc;

        public ControladorPrincipal()
        {
            cc = new CesantiaController();
            ac = new AsfamController();
            sc = new SILController();
            mc = new MaternalController();
        }

        public void ProcesarHistoricos(System.ComponentModel.BackgroundWorker worker, string periodo)
        {

            worker.ReportProgress(0);
            Fondos f = this.DatosExcelEstadisticas(periodo);
            worker.ReportProgress(40);
            
            cc.Procesar(f.Cesantia, periodo);
            worker.ReportProgress(60);
            ac.Procesar(f.Asfam, periodo);
            worker.ReportProgress(80);
            sc.Procesar(f.Sil, periodo);
            worker.ReportProgress(100);
            
        }


        public void ProcesarFondoCesantia(System.ComponentModel.BackgroundWorker worker, string periodo)
        {
            worker.ReportProgress(0);
            Cesantia c = new Cesantia()
            {
                AporteFiscalMes = Convert.ToString(DatosCuentasFBL3N("7003000002", periodo) * -1),
                Reintego = Convert.ToString(DatosCuentasFBL3N("7003000003", periodo) * -1),
                SubsidiosCesantia = Convert.ToString(DatosCuentasFBL3N("8007000001", periodo)),
                SubsidiosCesantiaRetroactivos = "0",
                ChequesCaducados="0",
                ChequesRevalidados="0",
                GastosDeAdministracion = "0"

            };
            worker.ReportProgress(70);

            cc.ProcesarFondo(c, periodo);
            worker.ReportProgress(100);
            

        }


        public void ProcesarFondoAsfam(System.ComponentModel.BackgroundWorker worker,string periodo)
        {
            worker.ReportProgress(0);

            List<long> Reintegros = new List<long>();
            Reintegros.Add(DatosCuentasFBL3N("7001000002", periodo));
            Reintegros.Add(DatosCuentasFBL3N("7001000007", periodo));
            Reintegros.Add(DatosCuentasFBL3N("7001000008", periodo));
            Reintegros.Add(DatosCuentasFBL3N("7001000009", periodo));

            Asfam c = new Asfam()
            {
                AporteFiscalMes = Convert.ToString(DatosCuentasFBL3N("7001000001", periodo) *-1),
                Reintego = Convert.ToString(Reintegros.Sum(i => i) * -1),
                AsFamTrabajadoresActivosMesActual = Convert.ToString(DatosCuentasFBL3N("8005000001", periodo)),
                AsFamPensionadosMesActual = "0",
                AsFamTrabajadoresCesantesMesActual = Convert.ToString(DatosCuentasFBL3N("8005000002", periodo)),
                AsFamInstitucionesMesActual = "0",

                AsFamTrabajadoresActivosRetroactivo = Convert.ToString(DatosCuentasFBL3N("8005000003", periodo)),
                AsFamPensionadosRetroactivo = "0",
                AsFamTrabajadoresCesantesRetroactivo = Convert.ToString(DatosCuentasFBL3N("8005000030", periodo)),
                AsFamInstitucionesRetroactivo = "0",

                DocumentosRevalidados = "0",
                ComisionAdministracion = Convert.ToString(DatosCuentasFBL3N("8005000008", periodo)),
                DocumentosCaducados = "0",
                DocumentosAnulados = "0",
                DevolucionDocumentosSAFEMCaducados = Convert.ToString(DatosCuentasFBL3N("8005000006", periodo) * -1),
                DevolucionDocumentosSAFEMAnulados = Convert.ToString(DatosCuentasFBL3N("8005000021", periodo) * -1),
                DocumentosSAFEMRevalidados = Convert.ToString(DatosCuentasFBL3N("8005000007", periodo)),

            };
            worker.ReportProgress(70);
            ac.ProcesarFondo(c, periodo);
            
            worker.ReportProgress(100);


        }


        public void ProcesarFondoSIL(System.ComponentModel.BackgroundWorker worker, string periodo)
        {
            worker.ReportProgress(0);

            List<long> CotizacionesPeriodosAnteriores = new List<long>();
            CotizacionesPeriodosAnteriores.Add(DatosCuentasFBL3N("7004000002", periodo));
            CotizacionesPeriodosAnteriores.Add(DatosCuentasFBL3N("7004000010", periodo));

            List<long> ReintegroporCobroIndebidodeSubsidio = new List<long>();
            ReintegroporCobroIndebidodeSubsidio.Add(DatosCuentasFBL3N("7004000003", periodo));
            ReintegroporCobroIndebidodeSubsidio.Add(DatosCuentasFBL3N("7004000007", periodo));

            List<long> SILEnfermedadOrigenComun = new List<long>();
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000001", periodo));
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000002", periodo));
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000010", periodo));

            
            List<long> SILSubsidioMaternalSuplementario = new List<long>();
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000011", periodo));
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000012", periodo));
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000013", periodo));

            List<long> DescuentoBeneficiosNoCobrados = new List<long>();
            DescuentoBeneficiosNoCobrados.Add(DatosCuentasFBL3N("8008000003", periodo));
            DescuentoBeneficiosNoCobrados.Add(DatosCuentasFBL3N("8008000020", periodo));

            List<long> SREnfermedadOrigenComun = new List<long>();
            SREnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000004", periodo));
            SREnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000019", periodo));
            

            SIL s = new SIL
            {
                Cotizaciones = Convert.ToString(DatosCuentasFBL3N("7004000009", periodo) * -1),
                CotizacionesPeriodosAnteriores = (CotizacionesPeriodosAnteriores.Sum(i => i) * -1).ToString(),
                ReajusteLey17332 = Convert.ToString(DatosCuentasFBL3N("7004000005", periodo) * -1),
                CotizacionesEntidadesPagadorasdeSubsidios = Convert.ToString(DatosCuentasFBL3N("7004000011", periodo) * -1),
                ReintegroporCobroIndebidodeSubsidio = (ReintegroporCobroIndebidodeSubsidio.Sum(i => i) * -1).ToString(),


                SILEnfermedadOrigenComun = SILEnfermedadOrigenComun.Sum(i => i).ToString(),
                SILSubsidioMaternalSuplementario = SILSubsidioMaternalSuplementario.Sum(i => i).ToString(),
                DescuentoBeneficiosNoCobrados = DescuentoBeneficiosNoCobrados.Sum(i => i).ToString(),

                SREnfermedadOrigenComun = SREnfermedadOrigenComun.Sum(i => i).ToString(),
                SRSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000008", periodo)),

                CFPEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000007", periodo)),
                CFPSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000014", periodo)),

                CFSEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000015", periodo)),
                CFSSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000016", periodo)),

                OCEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000017", periodo)),
                OCSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000018", periodo)),

                ComisionAdministracion = Convert.ToString(DatosCuentasFBL3N("8008000005", periodo)),
                OtrosEgresos = Convert.ToString(DatosCuentasFBL3N("8008000009", periodo)),

            };

            worker.ReportProgress(70);
            sc.ProcesarFondo(s, periodo);

            worker.ReportProgress(100);

        }


        public void ProcesarFondoMaternal(System.ComponentModel.BackgroundWorker worker, string periodo)
        {
            worker.ReportProgress(0);

            Maternal m = new Maternal
            {
                A1 = Convert.ToString(DatosCuentasFBL3N("7002000002", periodo) * -1),
                A2 = Convert.ToString(DatosCuentasFBL3N("7002000005", periodo) * -1),
                A31 = Convert.ToString(DatosCuentasFBL3N("7002000006", periodo) * -1),
                A32 = Convert.ToString(DatosCuentasFBL3N("7002000007", periodo) * -1),
                A41 = Convert.ToString(DatosCuentasFBL3N("7002000004", periodo) * -1),
                A42 = Convert.ToString(DatosCuentasFBL3N("7002000003", periodo) * -1),
                C1 = Convert.ToString(DatosCuentasFBL3N("8006000009", periodo)),
                C2 = Convert.ToString(DatosCuentasFBL3N("8006000010", periodo)),
                C3 = Convert.ToString(DatosCuentasFBL3N("8006000021", periodo)),
                C4 = Convert.ToString(DatosCuentasFBL3N("8006000011", periodo)),
                C5 = Convert.ToString(0),
                C61 = Convert.ToString(DatosCuentasFBL3N("8006000022", periodo)),
                C62 = Convert.ToString(DatosCuentasFBL3N("8006000023", periodo)),
                C63 = Convert.ToString(DatosCuentasFBL3N("8006000024", periodo)),
                C64 = Convert.ToString(DatosCuentasFBL3N("8006000025", periodo)),
                C65 = Convert.ToString(0),
                C71 = Convert.ToString(DatosCuentasFBL3N("8006000026", periodo)),
                C72 = Convert.ToString(DatosCuentasFBL3N("8006000027", periodo)),
                C73 = Convert.ToString(DatosCuentasFBL3N("8006000028", periodo)),
                C74 = Convert.ToString(DatosCuentasFBL3N("8006000029", periodo)),
                C75 = Convert.ToString(0),

                C81 = Convert.ToString(DatosCuentasFBL3N("8006000030", periodo)),
                C82 = Convert.ToString(DatosCuentasFBL3N("8006000031", periodo)),
                C83 = Convert.ToString(DatosCuentasFBL3N("8006000032", periodo)),
                C84 = Convert.ToString(DatosCuentasFBL3N("8006000033", periodo)),
                C85 = Convert.ToString(0),

                C91 = Convert.ToString(DatosCuentasFBL3N("8006000034", periodo)),
                C92 = Convert.ToString(DatosCuentasFBL3N("8006000035", periodo)),
                C93 = Convert.ToString(DatosCuentasFBL3N("8006000036", periodo)),
                C94 = Convert.ToString(DatosCuentasFBL3N("8006000037", periodo)),
                C95 = Convert.ToString(0),

                E1 = Convert.ToString(DatosCuentasFBL3N("8006000012", periodo)),
                E2 = Convert.ToString(DatosCuentasFBL3N("8006000013", periodo)),
                E3 = Convert.ToString(DatosCuentasFBL3N("8006000038", periodo)),
                E4 = Convert.ToString(DatosCuentasFBL3N("8006000014", periodo)),
                E5 = Convert.ToString(0),

                F1 = Convert.ToString(DatosCuentasFBL3N("8006000015", periodo)),
                F2 = Convert.ToString(DatosCuentasFBL3N("8006000016", periodo)),
                F3 = Convert.ToString(DatosCuentasFBL3N("8006000039", periodo)),
                F4 = Convert.ToString(DatosCuentasFBL3N("8006000017", periodo)),
                F5 = Convert.ToString(0),

                G1 = Convert.ToString(0),
                G2 = Convert.ToString(0),
                G3 = Convert.ToString(0),
                G4 = Convert.ToString(0),
                

            };

            worker.ReportProgress(80);
            mc.ProcesarFondo(m, periodo);

            worker.ReportProgress(100);
        }

        public void ProcesarAnexos(System.ComponentModel.BackgroundWorker worker, string periodo)
        {
            worker.ReportProgress(0);
            sc.GenerarAnexo(periodo);
            worker.ReportProgress(50);
            mc.GenerarAnexo(periodo);
            worker.ReportProgress(100);
        }


        public Fondos DatosExcelEstadisticas(string periodo)
        {

            Fondos fret = new Fondos();
            
            var excelApp = new ExcelX.Application();
            Utilidades.AbrirLibro(excelApp, @"C:\Fondos Nacionales\in\" + periodo + @"\EstadisticaCRF");
            

            #region Datos correspondientes al cuadro 1
            ExcelX._Worksheet Cuadro1 = (ExcelX.Worksheet)excelApp.Sheets["Cuadros N°1"];
            //Nº de Empresas
            fret.Cesantia.NroEmpresas  = Cuadro1.Range["I17"].Text;
            fret.Asfam.NroEmpresas = Cuadro1.Range["I17"].Text;
            fret.Sil.NumEmpresasAfiliadas = Cuadro1.Range["I17"].Text;
            fret.Sil.NumEmpresasCotizantes = Cuadro1.Range["I17"].Text;
            #endregion

            #region Datos correspondientes al cuadro 2
            ExcelX._Worksheet Cuadro2 = (ExcelX.Worksheet)excelApp.Sheets["Cuadros N°2"];
            //Nº de Afiliados en el mes anterior
            fret.Cesantia.NroAfiliados = Cuadro2.Range["G22"].Text;
            fret.Sil.NumTrabajadoresAfiliados = Cuadro2.Range["G22"].Text;
            #endregion

            #region Datos correspondientes al cuadro 3
            ExcelX._Worksheet Cuadro3 = (ExcelX.Worksheet)excelApp.Sheets["Cuadros N°3"];
            //Nº de Asignaciones Familiares Pagadas
            fret.Asfam.NroAfiliados = Cuadro3.Range["H32"].Text;
            #endregion

            #region Datos correspondientes al cuadro 7
            ExcelX._Worksheet Cuadro7 = (ExcelX.Worksheet)excelApp.Sheets["Cuadros N°7"];
            //Nº de Asignaciones Familiares Pagadas
            fret.Asfam.NroAsignacionesFamiliaresPagadas = Cuadro7.Range["I20"].Text;
            #endregion
            
            #region Datos correspondientes al cuadro 13
            ExcelX._Worksheet Cuadro13 = (ExcelX.Worksheet)excelApp.Sheets["Cuadro N°13"];
            //Nº de Subsidios Pagados en el mes anterior
            fret.Cesantia.NroSubsidios= Cuadro13.Range["F19"].Text;
            #endregion

            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);

            /*OTHR*/

            var excelAppXD = new ExcelX.Application();
            Utilidades.AbrirLibro(excelAppXD, @"C:\Fondos Nacionales\in\" + periodo + @"\EstadisticaED");
            
            ExcelX._Worksheet hoja = (ExcelX.Worksheet)excelAppXD.Sheets[1];
            int TtlSubsidios = Convert.ToInt32(hoja.Range["F10"].Text.Replace(".", "").Replace(",", ""));
            excelAppXD.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppXD);

            /*OTR*/
            var excelAppSil = new ExcelX.Application();
            Utilidades.AbrirLibro(excelAppSil, @"C:\Fondos Nacionales\in\" + periodo + @"\EstadisticaSIL");
            //excelAppSil.Workbooks.Open(@"C:\Fondos Nacionales\in\" + periodo + @"\EstadisticaSIL.xlsx");

            #region Datos correspondientes al cuadro 3
            ExcelX._Worksheet CuadroN3 = (ExcelX.Worksheet)excelAppSil.Sheets["SIL Anexo 6 - Cuadro Nº 3"];
            fret.Sil.NumSubsidiosIniciados = Convert.ToString(Convert.ToInt32(CuadroN3.Range["D34"].Text.Replace(".", "").Replace(",", "")) + Convert.ToInt32(CuadroN3.Range["E34"].Text.Replace(".", "").Replace(",", "")) + TtlSubsidios);
            #endregion

            #region Datos correspondientes al cuadro 6y7
            ExcelX._Worksheet CuadroN6Y7 = (ExcelX.Worksheet)excelAppSil.Sheets["SIL Anexo 6-Cuadro Nº 6 Y 7"];
            fret.Sil.NumAfiliadosCotizantes = CuadroN6Y7.Range["D32"].Text;
            #endregion

            excelAppSil.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppSil);



            /*Nota Interna*/

            var excelAppNotaInterna = new ExcelX.Application();
            var libroXXX = Utilidades.AbrirLibro(excelAppNotaInterna, @"C:\Fondos Nacionales\in\" + periodo + @"\NOTA_INTERNA");
            ExcelX._Worksheet HojaNotaInterna = libroXXX.Sheets["Nota Interna"];

            fret.Cesantia.ValorNotaInterna = HojaNotaInterna.Range["B10"].Text;
            fret.Sil.ValorNotaInterna = HojaNotaInterna.Range["B2"].Text;
            fret.Asfam.NI_Tramo0 = HojaNotaInterna.Range["E3"].Text;
            fret.Asfam.NI_Tramo1 = HojaNotaInterna.Range["E4"].Text;
            fret.Asfam.NI_Tramo2 = HojaNotaInterna.Range["E5"].Text;
            fret.Asfam.NI_Tramo3 = HojaNotaInterna.Range["E6"].Text;
            fret.Asfam.NI_Tramo4 = HojaNotaInterna.Range["E7"].Text;

            excelAppNotaInterna.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppNotaInterna);


            return fret;
        }


        public long DatosCuentasFBL3N(string cuenta, string periodo)
        {
            var excelApp = new ExcelX.Application();
            var Book = Utilidades.AbrirLibro(excelApp, @"C:\Fondos Nacionales\in\" + periodo + @"\FBL3N"); 
            ExcelX._Worksheet Cuentas = (ExcelX.Worksheet)excelApp.Sheets["Data"];
            var rango = Cuentas.UsedRange;

            rango.AutoFilter(Field: 1, Criteria1: cuenta);
            var result = rango.SpecialCells(ExcelX.XlCellType.xlCellTypeVisible, Type.Missing);

            List<long> a = new List<long>();

            foreach (ExcelX.Range row in result.Rows)
            {
                long correcto = 0;
                
                try
                {
                    var x = Convert.ToString(row.Cells[1, 2].Value2).Replace(".", "").Replace(",", "");
                    if (long.TryParse(x, out correcto))
                    {
                        a.Add(Convert.ToInt64(x));
                    }
                }
                catch (Exception es)
                {

                }
            }

            Book.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);

            long k = a.Sum(i => i);
            return k;
        }




    }

    public static class Utilidades
    {
        public static ExcelX.Workbook AbrirLibro(ExcelX.Application instancia, string ruta)
        {
            ExcelX.Workbook retorno;
            string extension = ".xlsx";

            try
            {
                retorno = instancia.Workbooks.Open(ruta + extension);
            }
            catch (Exception ex)
            {
                extension = ".xls";
                retorno = instancia.Workbooks.Open(ruta + extension);
            }

            return retorno;
        }

        public static string ExtensionLibro(ExcelX.Workbook libro)
        {
            return System.IO.Path.GetExtension(libro.FullName);
        }

        
    }
}
