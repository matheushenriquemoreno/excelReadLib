using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibExcel.Atributos
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
    public class IndexHeaderAttribute : Attribute
    {
        public int IndexColuna {  get; set; }
        public IndexHeaderAttribute(int indexColuna)
        {
            IndexColuna = indexColuna;
        }
    }
}
