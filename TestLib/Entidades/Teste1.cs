using LibExcel.Atributos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestLib.Entidades
{
    internal class Teste1
    {
        public class Pessoa
        {
            [IndexHeader(1)]
            public int Id { get; set; }
            public string Nome { get; set; }
            public int Idade { get; set; }

            [NomeHeader(NomeHeaderProperty = "Data De Nascimento")]
            public DateTime DataNascimento { get; set; }
        }
    }
}
