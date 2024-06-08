using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibExcel.Entidades.Exceptions
{
    public class ExcelNaoEncontrado : Exception
    {
        public ExcelNaoEncontrado(string mensage = "Caminho da planilha informada não foi encontrado!"): base(mensage)
        {
        }
    }
}
