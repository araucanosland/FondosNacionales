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

        public void ProcesarHistoricos(System.ComponentModel.BackgroundWorker worker)
        {

            worker.ReportProgress(0);
            Fondos f = this.DatosExcelEstadisticas();
            worker.ReportProgress(40);
            
            cc.Procesar(f.Cesantia);
            worker.ReportProgress(60);
            ac.Procesar(f.Asfam);
            worker.ReportProgress(80);
            sc.Procesar(f.Sil);
            worker.ReportProgress(100);
            
        }


        public void ProcesarFondoCesantia(System.ComponentModel.BackgroundWorker worker)
        {
            worker.ReportProgress(0);
            Cesantia c = new Cesantia()
            {
                AporteFiscalMes = Convert.ToString(DatosCuentasFBL3N("7003000002") * -1),
                Reintego = Convert.ToString(DatosCuentasFBL3N("7003000003") * -1),
                SubsidiosCesantia = Convert.ToString(DatosCuentasFBL3N("8007000001")),
                SubsidiosCesantiaRetroactivos = "0",
                ChequesCaducados="0",
                ChequesRevalidados="0",
                GastosDeAdministracion = "0"

            };
            worker.ReportProgress(70);

            cc.ProcesarFondo(c);
            worker.ReportProgress(100);
            

        }


        public void ProcesarFondoAsfam(System.ComponentModel.BackgroundWorker worker)
        {
            worker.ReportProgress(0);

            List<long> Reintegros = new List<long>();
            Reintegros.Add(DatosCuentasFBL3N("7001000002"));
            Reintegros.Add(DatosCuentasFBL3N("7001000007"));
            Reintegros.Add(DatosCuentasFBL3N("7001000008"));
            Reintegros.Add(DatosCuentasFBL3N("7001000009"));

            Asfam c = new Asfam()
            {
                AporteFiscalMes = Convert.ToString(DatosCuentasFBL3N("7001000001")*-1),
                Reintego = Convert.ToString(Reintegros.Sum(i => i) * -1),
                AsFamTrabajadoresActivosMesActual = Convert.ToString(DatosCuentasFBL3N("8005000001")),
                AsFamPensionadosMesActual = "0",
                AsFamTrabajadoresCesantesMesActual = Convert.ToString(DatosCuentasFBL3N("8005000002")),
                AsFamInstitucionesMesActual = "0",

                AsFamTrabajadoresActivosRetroactivo = Convert.ToString(DatosCuentasFBL3N("8005000003")),
                AsFamPensionadosRetroactivo = "0",
                AsFamTrabajadoresCesantesRetroactivo = Convert.ToString(DatosCuentasFBL3N("8005000030")),
                AsFamInstitucionesRetroactivo = "0",

                DocumentosRevalidados = "0",
                ComisionAdministracion = Convert.ToString(DatosCuentasFBL3N("8005000008")),
                DocumentosCaducados = "0",
                DocumentosAnulados = "0",
                DevolucionDocumentosSAFEMCaducados = Convert.ToString(DatosCuentasFBL3N("8005000006")* -1),
                DevolucionDocumentosSAFEMAnulados = Convert.ToString(DatosCuentasFBL3N("8005000021") * -1),
                DocumentosSAFEMRevalidados = Convert.ToString(DatosCuentasFBL3N("8005000007")),

            };
            worker.ReportProgress(70);
            ac.ProcesarFondo(c);
            
            worker.ReportProgress(100);


        }


        public void ProcesarFondoSIL(System.ComponentModel.BackgroundWorker worker)
        {
            worker.ReportProgress(0);

            List<long> CotizacionesPeriodosAnteriores = new List<long>();
            CotizacionesPeriodosAnteriores.Add(DatosCuentasFBL3N("7004000002"));
            CotizacionesPeriodosAnteriores.Add(DatosCuentasFBL3N("7004000010"));

            List<long> ReintegroporCobroIndebidodeSubsidio = new List<long>();
            ReintegroporCobroIndebidodeSubsidio.Add(DatosCuentasFBL3N("7004000003"));
            ReintegroporCobroIndebidodeSubsidio.Add(DatosCuentasFBL3N("7004000007"));

            List<long> SILEnfermedadOrigenComun = new List<long>();
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000001"));
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000002"));
            SILEnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000010"));

            
            List<long> SILSubsidioMaternalSuplementario = new List<long>();
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000011"));
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000012"));
            SILSubsidioMaternalSuplementario.Add(DatosCuentasFBL3N("8008000013"));

            List<long> DescuentoBeneficiosNoCobrados = new List<long>();
            DescuentoBeneficiosNoCobrados.Add(DatosCuentasFBL3N("8008000003"));
            DescuentoBeneficiosNoCobrados.Add(DatosCuentasFBL3N("8008000020"));

            List<long> SREnfermedadOrigenComun = new List<long>();
            SREnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000004"));
            SREnfermedadOrigenComun.Add(DatosCuentasFBL3N("8008000019"));
            

            SIL s = new SIL
            {
                Cotizaciones = Convert.ToString(DatosCuentasFBL3N("7004000009") * -1),
                CotizacionesPeriodosAnteriores = (CotizacionesPeriodosAnteriores.Sum(i => i) * -1).ToString(),
                ReajusteLey17332 = Convert.ToString(DatosCuentasFBL3N("7004000005") * -1),
                CotizacionesEntidadesPagadorasdeSubsidios = Convert.ToString(DatosCuentasFBL3N("7004000011") * -1),
                ReintegroporCobroIndebidodeSubsidio = (ReintegroporCobroIndebidodeSubsidio.Sum(i => i) * -1).ToString(),


                SILEnfermedadOrigenComun = SILEnfermedadOrigenComun.Sum(i => i).ToString(),
                SILSubsidioMaternalSuplementario = SILSubsidioMaternalSuplementario.Sum(i => i).ToString(),
                DescuentoBeneficiosNoCobrados = DescuentoBeneficiosNoCobrados.Sum(i => i).ToString(),

                SREnfermedadOrigenComun = SREnfermedadOrigenComun.Sum(i => i).ToString(),
                SRSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000008")),

                CFPEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000007")),
                CFPSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000014")),

                CFSEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000015")),
                CFSSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000016")),

                OCEnfermedadOrigenComun = Convert.ToString(DatosCuentasFBL3N("8008000017")),
                OCSubsidioMaternalSuplementario = Convert.ToString(DatosCuentasFBL3N("8008000018")),

                ComisionAdministracion = Convert.ToString(DatosCuentasFBL3N("8008000005")),
                OtrosEgresos = Convert.ToString(DatosCuentasFBL3N("8008000009")),

            };

            worker.ReportProgress(70);
            sc.ProcesarFondo(s);

            worker.ReportProgress(100);

        }


        public void ProcesarFondoMaternal(System.ComponentModel.BackgroundWorker worker)
        {
            worker.ReportProgress(0);

            Maternal m = new Maternal
            {
                A1 = Convert.ToString(DatosCuentasFBL3N("7002000002") * -1),
                A2 = Convert.ToString(DatosCuentasFBL3N("7002000005") * -1),
                A31 = Convert.ToString(DatosCuentasFBL3N("7002000006") * -1),
                A32 = Convert.ToString(DatosCuentasFBL3N("7002000007") * -1),
                A41 = Convert.ToString(DatosCuentasFBL3N("7002000004") * -1),
                A42 = Convert.ToString(DatosCuentasFBL3N("7002000003") * -1),
                C1 = Convert.ToString(DatosCuentasFBL3N("8006000009")),
                C2 = Convert.ToString(DatosCuentasFBL3N("8006000010")),
                C3 = Convert.ToString(DatosCuentasFBL3N("8006000021")),
                C4 = Convert.ToString(DatosCuentasFBL3N("8006000011")),
                C5 = Convert.ToString(0),
                C61 = Convert.ToString(DatosCuentasFBL3N("8006000022")),
                C62 = Convert.ToString(DatosCuentasFBL3N("8006000023")),
                C63 = Convert.ToString(DatosCuentasFBL3N("8006000024")),
                C64 = Convert.ToString(DatosCuentasFBL3N("8006000025")),
                C65 = Convert.ToString(0),
                C71 = Convert.ToString(DatosCuentasFBL3N("8006000026")),
                C72 = Convert.ToString(DatosCuentasFBL3N("8006000027")),
                C73 = Convert.ToString(DatosCuentasFBL3N("8006000028")),
                C74 = Convert.ToString(DatosCuentasFBL3N("8006000029")),
                C75 = Convert.ToString(0),

                C81 = Convert.ToString(DatosCuentasFBL3N("8006000030")),
                C82 = Convert.ToString(DatosCuentasFBL3N("8006000031")),
                C83 = Convert.ToString(DatosCuentasFBL3N("8006000032")),
                C84 = Convert.ToString(DatosCuentasFBL3N("8006000033")),
                C85 = Convert.ToString(0),

                C91 = Convert.ToString(DatosCuentasFBL3N("8006000034")),
                C92 = Convert.ToString(DatosCuentasFBL3N("8006000035")),
                C93 = Convert.ToString(DatosCuentasFBL3N("8006000036")),
                C94 = Convert.ToString(DatosCuentasFBL3N("8006000037")),
                C95 = Convert.ToString(0),

                E1 = Convert.ToString(DatosCuentasFBL3N("8006000012")),
                E2 = Convert.ToString(DatosCuentasFBL3N("8006000013")),
                E3 = Convert.ToString(DatosCuentasFBL3N("8006000038")),
                E4 = Convert.ToString(DatosCuentasFBL3N("8006000014")),
                E5 = Convert.ToString(0),

                F1 = Convert.ToString(DatosCuentasFBL3N("8006000015")),
                F2 = Convert.ToString(DatosCuentasFBL3N("8006000016")),
                F3 = Convert.ToString(DatosCuentasFBL3N("8006000039")),
                F4 = Convert.ToString(DatosCuentasFBL3N("8006000017")),
                F5 = Convert.ToString(0),

                G1 = Convert.ToString(0),
                G2 = Convert.ToString(0),
                G3 = Convert.ToString(0),
                G4 = Convert.ToString(0),
                

            };

            worker.ReportProgress(80);
            mc.ProcesarFondo(m);

            worker.ReportProgress(100);
        }

        public void ProcesarAnexos(System.ComponentModel.BackgroundWorker worker)
        {
            worker.ReportProgress(0);
            sc.GenerarAnexo();
            worker.ReportProgress(100);
        }


        public Fondos DatosExcelEstadisticas()
        {

            Fondos fret = new Fondos();
            
            var excelApp = new ExcelX.Application();
            excelApp.Workbooks.Open(@"C:\Fondos Nacionales\in\Est201701CRF.xls");

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
            excelAppXD.Workbooks.Open(@"C:\Fondos Nacionales\in\Est201701ED.xls");
            ExcelX._Worksheet hoja = (ExcelX.Worksheet)excelAppXD.Sheets[1];
            int TtlSubsidios = Convert.ToInt32(hoja.Range["F10"].Text.Replace(".", "").Replace(",", ""));
            excelAppXD.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelAppXD);

            /*OTR*/
            var excelAppSil = new ExcelX.Application();
            excelAppSil.Workbooks.Open(@"C:\Fondos Nacionales\in\Est201701SIL.xlsx");

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

            return fret;
        }


        public long DatosCuentasFBL3N(string cuenta)
        {
            var excelApp = new ExcelX.Application();
            var Book = excelApp.Workbooks.Open(@"C:\Fondos Nacionales\in\201701\201701FBL3N.xlsx");
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
}
