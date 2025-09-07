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

        if (!arquivo.Exists)
        {
            using (var pkgNovo = new ExcelPackage())
            {
                pkgNovo.Workbook.Worksheets.Add("Sheet1");
                pkgNovo.SaveAs(arquivo);
            }
        }
        using (var package = new ExcelPackage(arquivo))
        {
            var worksheets = package.Workbook.Worksheets;
            var idxPrimeiro = package.Compatibility.IsWorksheets1Based ? 1 : 0;
            var Informacoes = new[]
           {
                new { Nome = "Daniel", Sobrenome="Galleazzo", Idade = 19 },
                new { Nome = "Paulo", Sobrenome="Galleazzo", Idade = 21 },
                new { Nome = "Júlia", Sobrenome="Zanon", Idade = 20 },
                new { Nome = "Sandra", Sobrenome="Galleazzo", Idade = 51 },
                new { Nome = "Antônio", Sobrenome="Galleazzo", Idade = 81 },
                new { Nome = "Maria", Sobrenome="Galleazzo", Idade = 80 },
                new { Nome = "Raissa", Sobrenome ="AnticristoSDD" , Idade = 666}
            };

            var planilha = worksheets.Count == 0
                ? worksheets.Add("Sheet1")
                : worksheets[idxPrimeiro];
            while (true)
            {
                Console.Write("Digite algo para adicionar (ou 'sair'): ");
                string entrada = Console.ReadLine();
                if (entrada == null) continue;

                if (entrada.Trim().ToLower() == "sair") break;

                planilha.Cells[1, 1].Value = "Nome"; //cabeçalho da planilha
                planilha.Cells[1, 2].Value = "Sobrenome";
                planilha.Cells[1, 3].Value = "Idade";
                planilha.Cells["A1:C1"].Style.Font.Bold = true;
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
        }
        Console.WriteLine($"Arquivo salvo em: {Path.GetFullPath(caminhoArquivo)}");
    }
}
