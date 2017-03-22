using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IA.FondosNacionales.Entity
{
    public class Fondos
    {
        public Cesantia Cesantia { get; set; }
        //public Maternal Maternal { get; set; }
        public Asfam Asfam { get; set; }
        public SIL Sil { get; set; }


        public Fondos()
        {
            this.Cesantia = new Cesantia();
            this.Asfam = new Asfam();
            this.Sil = new SIL();
        }
    }
}
