using LibExcel.Atributos;
using LibExcel.Entidades;
using LibExcel.Entidades.Exceptions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static OfficeOpenXml.ExcelErrorValue;

namespace LibExcel.Services
{
    public class LeitorExcell<T> where T : new()
    {

        public int PosicaoPlanilha { get; set; } = 0;
        public int LinhaHeader { get; set; } = 1;

        public LeitorExcell()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private (ExcelWorksheet excell, List<ColunasExcell> colunas) ObterDadosHeaderExcel(string caminhoPlanilha)
        {
            if (!File.Exists(caminhoPlanilha))
                throw new ExcelNaoEncontrado();

            var package = new ExcelPackage(new FileInfo(caminhoPlanilha));

            ExcelWorksheet worksheet = package.Workbook.Worksheets[PosicaoPlanilha];

            var colunasHeader = new List<ColunasExcell>();

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string campo = worksheet.Cells[LinhaHeader, col].Text;

                if (string.IsNullOrEmpty(campo))
                    break;

                colunasHeader.Add(new ColunasExcell(campo, col));
            }

            return(worksheet, colunasHeader);
        }
    
        public List<T> ObterDadosDoExcel(string caminhoPlanilha)
        {
            var dados = new List<T>();

            var (worksheet, colunasHeader) = ObterDadosHeaderExcel(caminhoPlanilha);

            int linhaDados = LinhaHeader + 1;
            PropertyInfo[] campos = typeof(T).GetProperties();

            for (int row = linhaDados; row <= worksheet.Dimension.End.Row; row++)
            {
                var objeto = new T();
                
                foreach (var coluna in colunasHeader)
                {
                    string valor = worksheet.Cells[row, coluna.Index].Text;

                    PreencherObjeto(campos, coluna, objeto, valor);
                }

                dados.Add(objeto);
            }

            worksheet.Dispose();

            return dados;
        }


        private void PreencherObjeto(PropertyInfo[] campos, ColunasExcell coluna, object entidade, object value)
        {
         
            foreach (PropertyInfo campo in campos)
            {
                if (!campo.CanWrite)
                    continue;

                if (PossuiAtributoIndex(campo) && EhCampoIndex(campo, coluna.Index))
                {
                    PreencherObjeto(campo, entidade, value);
                    break;
                }

                if (PossuiAtributoNomeHeader(campo) && EhCampoNomePersonalizado(campo, coluna.Nome))
                {
                    PreencherObjeto(campo, entidade, value);
                    break;
                }

                if (coluna.NomeIgual(campo.Name))
                {
                    PreencherObjeto(campo, entidade, value);
                    break;
                }
            }
        }

        private bool PossuiAtributoIndex(PropertyInfo campo)
        {
            var result = (IndexHeaderAttribute)campo.GetCustomAttributes(typeof(IndexHeaderAttribute), false).FirstOrDefault();

            return result is not null ? true : false;
        }

        private bool EhCampoIndex(PropertyInfo campo, int index)
        {
            var result = (IndexHeaderAttribute)campo.GetCustomAttributes(typeof(IndexHeaderAttribute), false).FirstOrDefault();

            return result.IndexColuna == index;
        }

        public void PreencherObjeto(PropertyInfo campo, object entidade, object value)
        {
            if (campo.PropertyType == typeof(int))
            {
                campo.SetValue(entidade, ConverterParaTipo<int>(value));
            }
            else if (campo.PropertyType == typeof(string))
            {
                campo.SetValue(entidade, ConverterParaTipo<string>(value));
            }
            else if (campo.PropertyType == typeof(DateTime))
            {
                if (DateTime.TryParse(value.ToString(), out DateTime valorConvertido))
                {
                    campo.SetValue(entidade, valorConvertido);
                }
            }
        }

        private bool PossuiAtributoNomeHeader(PropertyInfo campo)
        {
            var result = (NomeHeaderAttribute)campo.GetCustomAttributes(typeof(NomeHeaderAttribute), false).FirstOrDefault();

            return result is not null ? true : false;
        }

        private bool EhCampoNomePersonalizado(PropertyInfo campo, string nomeColuna)
        {
            var result = (NomeHeaderAttribute)campo.GetCustomAttributes(typeof(NomeHeaderAttribute), false).FirstOrDefault();

            return result.NomeHeaderProperty.Equals(nomeColuna, StringComparison.CurrentCultureIgnoreCase);
        }

        static Treturn ConverterParaTipo<Treturn>(object valorString)
        {
            TypeConverter converter = TypeDescriptor.GetConverter(typeof(Treturn));
            return (Treturn)converter.ConvertFrom(valorString);
        }

    }
}
