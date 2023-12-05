using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibExcel.Entidades
{
    internal class ColunasExcell
    {
        public string Nome { get; set; }
        public int Index {  get; set; }

        public ColunasExcell(string nomeColuna, int indexColuna)
        {
            Nome = nomeColuna;
            Index = indexColuna;
        }


        public bool NomeIgual(string nome)
        {
            return this.Nome.Equals(nome, StringComparison.CurrentCultureIgnoreCase);
        }

    }
}
