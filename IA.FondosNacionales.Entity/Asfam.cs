using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IA.FondosNacionales.Entity
{
    public class Asfam
    {
        public string NroAsignacionesFamiliaresPagadas { get; set; }
        public string NroAfiliados { get; set; }
        public string NroEmpresas { get; set; }

        public string AporteFiscalMes { get; set; }
        public string Reintego { get; set; }

        public string AsFamTrabajadoresActivosMesActual { get; set; }
        public string AsFamPensionadosMesActual { get; set; }
        public string AsFamTrabajadoresCesantesMesActual { get; set; }
        public string AsFamInstitucionesMesActual { get; set; }

        public string AsFamTrabajadoresActivosRetroactivo { get; set; }
        public string AsFamPensionadosRetroactivo { get; set; }
        public string AsFamTrabajadoresCesantesRetroactivo { get; set; }
        public string AsFamInstitucionesRetroactivo { get; set; }

        public string DocumentosRevalidados { get; set; }
        public string ComisionAdministracion { get; set; }
        public string DocumentosCaducados { get; set; }
        public string DocumentosAnulados { get; set; }
        public string DevolucionDocumentosSAFEMCaducados { get; set; }
        public string DevolucionDocumentosSAFEMAnulados { get; set; }
        public string DocumentosSAFEMRevalidados { get; set; }
    }
}
