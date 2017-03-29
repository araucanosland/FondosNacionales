using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IA.FondosNacionales.Entity
{
    public class SIL
    {
        public string NumSubsidiosIniciados { get; set; }
        public string NumAfiliadosCotizantes { get; set; }
        public string NumEmpresasCotizantes { get; set; }
        public string NumTrabajadoresAfiliados { get; set; }
        public string NumEmpresasAfiliadas { get; set; }


        public string Cotizaciones { get; set; }
        public string CotizacionesPeriodosAnteriores { get; set; }
        public string ReajusteLey17332 { get; set; }
        public string CotizacionesEntidadesPagadorasdeSubsidios { get; set; }
        public string ReintegroporCobroIndebidodeSubsidio { get; set; }


        public string SILEnfermedadOrigenComun { get; set; }
        public string SILSubsidioMaternalSuplementario { get; set; }

        public string DescuentoBeneficiosNoCobrados { get; set; }
        
        public string SREnfermedadOrigenComun { get; set; }
        public string SRSubsidioMaternalSuplementario { get; set; }

        public string CFPEnfermedadOrigenComun { get; set; }
        public string CFPSubsidioMaternalSuplementario { get; set; }

        public string CFSEnfermedadOrigenComun { get; set; }
        public string CFSSubsidioMaternalSuplementario { get; set; }

        public string OCEnfermedadOrigenComun { get; set; }
        public string OCSubsidioMaternalSuplementario { get; set; }


        public string ComisionAdministracion { get; set; }
        public string OtrosEgresos { get; set; }


        public string ValorNotaInterna { get; set; }

    }
}
