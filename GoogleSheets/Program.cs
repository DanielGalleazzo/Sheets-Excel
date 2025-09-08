using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string caminhoArquivo = @"";
        var arquivo = new FileInfo(caminhoArquivo);
        ExcelPackage.License.SetNonCommercialPersonal("DanielGalleazzo");

     

        using (var package = new ExcelPackage(arquivo))
        {
            var worksheets = package.Workbook.Worksheets;
            var idxPrimeiro = package.Compatibility.IsWorksheets1Based ? 1 : 0;
            var Informacoes = new[]
           {
                new { Nome = "Itqachi", Sobrenome="Galleazzo", Idade = 19 },
                new { Nome = "sasuke", Sobrenome="Galleazzo", Idade = 21 },
                new { Nome = "oii", Sobrenome="Zanon", Idade = 20 },
                new { Nome = "Receba", Sobrenome="Galleazzo", Idade = 51 },
                new { Nome = "Luva ", Sobrenome="Galleazzo", Idade = 81 },
                new { Nome = "de pedreiro", Sobrenome="Galleazzo", Idade = 80 },
                new { Nome = "Gabriel", Sobrenome ="Testando" , Idade = 666}
            };
            
            var planilha = worksheets.Count == 0
                ? worksheets.Add("Sheet1")
                : worksheets[idxPrimeiro];

           
            
                Console.WriteLine("Você tem certeza que quer adicionar esses valores ? ( sim ou não ) ");
                string entrada = Console.ReadLine();

                if (entrada == "sim")
                {
                    planilha.Cells[1, 1].Value = "Nome"; //cabeçalho da planilha
                    planilha.Cells[1, 2].Value = "Sobrenome";
                    planilha.Cells[1, 3].Value = "Idade";


                    planilha.Cells["A1:C1"].Style.Font.Bold = true; // estilo da fonte
                    planilha.Cells["A1:C1"].Style.Font.Italic = true;

                    int indice = 2;
                    foreach (var informacoes in Informacoes)
                    {
                        planilha.Cells[indice, 1].Value = informacoes.Nome;
                        planilha.Cells[indice, 2].Value = informacoes.Sobrenome;
                        planilha.Cells[indice, 3].Value = informacoes.Idade;
                        indice++;
                    }

                    package.Save();
                    Console.WriteLine("Adicionado!");
                }

                if (entrada.Trim().ToLower() == "não")
                {
                    Console.WriteLine("Volte quando tiver certeza");
                   
                }


               
            
        }
        Console.WriteLine($"Arquivo salvo em: {Path.GetFullPath(caminhoArquivo)}");
    }
}