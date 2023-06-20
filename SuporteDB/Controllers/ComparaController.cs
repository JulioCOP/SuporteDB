using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using SuporteDB.Models;
using System.Collections.Generic;
using System.IO;

[Route("api/[controller]")]
[ApiController]
public class ComparacaoController : ControllerBase
{
    private readonly IWebHostEnvironment _hostingEnvironment;

    public ComparacaoController(IWebHostEnvironment hostingEnvironment)
    {
        _hostingEnvironment = hostingEnvironment;
    }

    [HttpPost("comparar")]
    public IActionResult CompararValores(List<ComparaDados> comparaDados)
    {
        // Salvar o arquivo recebido
        var file = Request.Form.Files[0];
        var filePath = Path.Combine(_hostingEnvironment.ContentRootPath, "uploads", file.FileName);

        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            file.CopyTo(stream);
        }

        // Fazer a leitura do arquivo e comparar os valores
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var result = new List<ComparaDados>();

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var valor1 = worksheet.Cells[row, 1].Value?.ToString();
                var valor2 = worksheet.Cells[row, 2].Value?.ToString();

                if (!string.IsNullOrEmpty(valor1) && !string.IsNullOrEmpty(valor2))
                {
                    var dados = new ComparaDados
                    {
                        Planilha1 = valor1,
                        Planilha2 = valor2
                    };

                    result.Add(dados);
                }
            }

            // Comparar os valores com os dados recebidos
            var comparacoes = new List<ComparaDados>();

            foreach (var dados in comparaDados)
            {
                var comparacao = result.FirstOrDefault(r => r.Planilha1 == dados.Planilha1 && r.Planilha2 == dados.Planilha2);

                if (comparacao != null)
                {
                    comparacoes.Add(comparacao);
                }
            }

            return Ok(comparacoes);
        }
    }
}
