using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class Program
{
    public static void Main(string[] args)
    {
        string caminhoPlanilha = @"";
        Console.WriteLine("Pressione algo para iniciar");
        Console.ReadKey();

        CriaPlanilhaExcel(caminhoPlanilha);

        Console.WriteLine("Pressione algo para acessar e exibir a planilha\n");
        Console.ReadKey();

        AbrePlanilhaExcel(caminhoPlanilha);

        Console.ReadKey();
    }

    static void CriaPlanilhaExcel(string caminhoPlanilha)
    {
        var Vendas = new[]
        {
            new { Id = "SP101", Filial="São Paulo", Vendas = 980 },
            new { Id = "RJ102", Filial="Rio de Janeiro", Vendas = 840 },
            new { Id = "MG103", Filial="Minas Gerais", Vendas = 790 },
            new { Id = "BA104", Filial="Bahia", Vendas = 699 },
            new { Id = "PR105", Filial="Paraná", Vendas = 775 },
            new { Id = "RS106", Filial="Porto Alegre", Vendas = 660 }
        };

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var excel = new ExcelPackage())
        {
            var workSheet = excel.Workbook.Worksheets.Add("PlanilhaVendas");

            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            workSheet.Cells[1, 1].Value = "Cod.";
            workSheet.Cells[1, 2].Value = "Filial";
            workSheet.Cells[1, 3].Value = "Vendas/mil";
            workSheet.Cells["A1:C1"].Style.Font.Italic = true;

            int indice = 2;
            foreach (var venda in Vendas)
            {
                workSheet.Cells[indice, 1].Value = venda.Id;
                workSheet.Cells[indice, 2].Value = venda.Filial;
                workSheet.Cells[indice, 3].Value = venda.Vendas;
                indice++;
            }

            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            if (File.Exists(caminhoPlanilha))
                File.Delete(caminhoPlanilha);

            File.WriteAllBytes(caminhoPlanilha, excel.GetAsByteArray());
        }

        Console.WriteLine($"Planilha criada com sucesso em : {caminhoPlanilha}\n");
    }

    static void AbrePlanilhaExcel(string caminhoPlanilha)
    {
        if (File.Exists(caminhoPlanilha))
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = caminhoPlanilha,
                UseShellExecute = true
            });
        }
        else
        {
            Console.WriteLine("Arquivo não encontrado.");
        }
    }
}
