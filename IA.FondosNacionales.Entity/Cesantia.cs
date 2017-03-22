using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IA.FondosNacionales.Entity
{
    public class Cesantia
    {
        public string NroEmpresas { get; set; }
        public string NroAfiliados { get; set; }
        public string NroSubsidios { get; set; }

        public string AporteFiscalMes { get; set; }
        public string Reintego { get; set; }
        public string SubsidiosCesantia { get; set; }
        public string SubsidiosCesantiaRetroactivos { get; set; }
        public string ChequesCaducados { get; set; }
        public string ChequesRevalidados { get; set; }
        public string GastosDeAdministracion { get; set; }

    }


}
