using LibExcel.Services;
using static TestLib.Entidades.Teste1;

var leitor = new LeitorExcell<Pessoa>();

//var caminhoPlanilha = @"C:\Users\mathe\Documents\ExcellTestes\Teste1.xlsx";
var caminhoPlanilha = @"C:\Users\mathe\Documents\ExcellTestes\Teste2.xlsx";


var pessoas = leitor.ObterDadosDoExcel(caminhoPlanilha);


foreach(var pessoa in pessoas)
{
    Console.WriteLine($"{pessoa.Nome}, {pessoa.Idade}, {pessoa.DataNascimento}, {pessoa.Id}");
}

Console.ReadLine();
