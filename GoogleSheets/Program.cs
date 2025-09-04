using OfficeOpenXml;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        string caminhoArquivo = @"adicione o caminho do seu arquivo";
        FileInfo arquivo = new FileInfo(caminhoArquivo);

       
        using (ExcelPackage package = new ExcelPackage(arquivo))
        {
            var planilha = package.Workbook.Worksheets.Count == 0
                ? package.Workbook.Worksheets.Add("Sheet1")
                : package.Workbook.Worksheets[0];

            while (true)
            {
                Console.Write("Digite algo para adicionar (ou 'sair'): ");
                string entrada = Console.ReadLine();

                if (entrada.ToLower() == "sair") break;

                int novaLinha = planilha.Dimension?.Rows + 1 ?? 1;
                planilha.Cells[novaLinha, 1].Value = entrada;
                planilha.Cells[novaLinha, 2].Value = DateTime.Now.ToString(); // momento exato em que foi adicionado

                package.Save();
                Console.WriteLine(" Adicionado");
            }
        }

        Console.WriteLine($"Arquivo salvo em: {Path.GetFullPath(caminhoArquivo)}");
    }
}
