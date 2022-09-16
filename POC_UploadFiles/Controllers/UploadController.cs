using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace POC_UploadFiles.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class UploadController : ControllerBase
    {
        //[HttpPost(Name = "upload")]
        //[RequestSizeLimit(bytes: 999_999_999_999_999)]
        //public async Task uploadPDFAsync(IEnumerable<IFormFile> files)
        //{
        //    var stopwatch = new Stopwatch();

        //    stopwatch.Start();

        //    foreach (var file in files)
        //    {

        //        Console.WriteLine(file.FileName);
        //        if (file.Length <= 0)
        //            return;


        //        //Strip out any path specifiers (ex: /../)
        //        var originalFileName = Path.GetFileName(file.FileName);


        //        //Create a unique file path
        //        var uniqueFileName = Path.GetRandomFileName();
        //        var uniqueFilePath = Path.Combine(@"C:\temp\", originalFileName);


        //        //Save the file to disk
        //        using (var stream = System.IO.File.Create(uniqueFilePath))
        //        {
        //            await file.CopyToAsync(stream);
        //        }
        //    }

        //    stopwatch.Stop();
        //    Console.WriteLine($"Tempo passado: {stopwatch.Elapsed}");

        //}


        [HttpPost]
        [RequestSizeLimit(bytes: 999_999_999_999_999)]
        public async Task sheetActiveAsync(IFormFile file)
        {

            //Strip out any path specifiers (ex: /../)
            var originalFileName = Path.GetFileName(file.FileName);


            //Create a unique file path
            var uniqueFileName = Path.GetRandomFileName();
            var uniqueFilePath = Path.Combine(@"C:\temp\", originalFileName);


            //Save the file to disk
            using (var stream = System.IO.File.Create(uniqueFilePath))
            {
                await file.CopyToAsync(stream);
            }
            var xls = new XLWorkbook(uniqueFilePath);
            var planilha = xls.Worksheets.First();
            var totalLinhas = planilha.Rows().Count();
            Console.WriteLine($"totalLinhas - {totalLinhas}");
            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                var OBRA_ID = planilha.Cell($"A{l}").Value.ToString();
                var NOME_OPERACAOO = planilha.Cell($"B{l}").Value.ToString();
                var EMPREENDIMENTO = planilha.Cell($"C{l}").Value.ToString();
                var IDENTIFICADOR = planilha.Cell($"D{l}").Value.ToString();
                var CLIENTES = planilha.Cell($"E{l}").Value.ToString();
                var CPF_CNPJ = planilha.Cell($"F{l}").Value.ToString();
                var DATA_VENDA = planilha.Cell($"G{l}").Value.ToString();
                var DATA_CESSAO = planilha.Cell($"H{l}").Value.ToString();
                var VALOR_VENDA = planilha.Cell($"I{l}").Value.ToString();


                if (OBRA_ID != "")
                {
                    Console.WriteLine($"{OBRA_ID} - {NOME_OPERACAOO} - {EMPREENDIMENTO} - {IDENTIFICADOR} - {CLIENTES} - {CPF_CNPJ} - {DATA_VENDA} - {DATA_CESSAO} - {VALOR_VENDA}");

                }
            }
        }
    }
}
