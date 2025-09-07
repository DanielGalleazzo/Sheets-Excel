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

            var planilha = worksheets.Count == 0
                ? worksheets.Add("Sheet1")
                : worksheets[idxPrimeiro];
            while (true)
            {
                Console.Write("Digite algo para adicionar (ou 'sair'): ");
                string entrada = Console.ReadLine();
                if (entrada == null) continue;

                if (entrada.Trim().ToLower() == "sair") break;

                int novaLinha = (planilha.Dimension?.Rows ?? 0) + 1;
                planilha.Cells[novaLinha, 1].Value = entrada;
                planilha.Cells[novaLinha, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                package.Save();
                Console.WriteLine("Adicionado!");
            }
        }
        Console.WriteLine($"Arquivo salvo em: {Path.GetFullPath(caminhoArquivo)}");
    }
}
