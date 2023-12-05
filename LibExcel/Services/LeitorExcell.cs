using LibExcel.Atributos;
using LibExcel.Entidades;
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
        public LeitorExcell()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public List<T> ObterDadosDoExcel(string caminhoPlanilha, int posicaoPlanilha = 0)
        {
            var dados = new List<T>();

            using var package = new ExcelPackage(new FileInfo(caminhoPlanilha));

            ExcelWorksheet worksheet = package.Workbook.Worksheets[posicaoPlanilha];

            var colunasHeader = new List<ColunasExcell>();

            // A primeira linha contém os nomes dos campos
            int rowCampo = 1;
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string campo = worksheet.Cells[rowCampo, col].Text;

                if (string.IsNullOrEmpty(campo))
                    break;

                colunasHeader.Add(new ColunasExcell(campo, col));
            }

            int linhaDados = rowCampo + 1;
            for (int row = linhaDados; row <= worksheet.Dimension.End.Row; row++)
            {
                var objeto = new T();

                // Preencher os valores do objeto com base nos nomes dos campos
                foreach (var coluna in colunasHeader)
                {
                    string valor = worksheet.Cells[row, coluna.Index].Text;

                    PreencherObjeto(coluna, objeto, valor);
                }

                dados.Add(objeto);
            }

            return dados;
        }


        private void PreencherObjeto(ColunasExcell coluna, object entidade, object value)
        {
            var campos = typeof(T).GetProperties();

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
            throw new NotImplementedException();
        }

        private bool EhCampoIndex(PropertyInfo campo, int index)
        {
            throw new NotImplementedException();
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
