using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibExcel.Atributos
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public class NomeHeaderAttribute : Attribute
    {
        /// <summary>
        /// Caso o nome do Header venha com caracter especial, ou com espaço, coloque o respequitivo nome aqui que o sistema ira fazer a comparação
        /// caso o nome seja igual o valor sera atribuido a classe.
        /// Não importa o padrão as string vão ser comparadas no minusculo
        /// </summary>
        public string NomeHeaderProperty { get; set; }
    }
}
