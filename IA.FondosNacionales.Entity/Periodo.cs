using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IA.FondosNacionales.Entity
{
    public class Periodo
    {
        public int Valor { get; set; }
        public string Texto { get; set; }

        public Periodo(int valor, string texto)
        {
            this.Valor = valor;
            this.Texto = Texto;

            
        }
        
        public override string ToString()
        {
            return this.Texto;
        }
    }
}
