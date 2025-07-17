using Microsoft.AspNetCore.Mvc;
using SlideMergerAPINew.Models;
using SlideMergerAPINew.Services;

namespace SlideMergerAPINew.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SlideMergerController : ControllerBase
    {
        private readonly SlideMergerService _slideMergerService;

        public SlideMergerController(SlideMergerService slideMergerService)
        {
            _slideMergerService = slideMergerService;
        }

        [HttpPost("merge")]
        public async Task<IActionResult> MergeSlides([FromForm] SlideMergeRequest request, IFormFile destinationFile)
        {
            if (destinationFile == null || destinationFile.Length == 0)
            {
                return BadRequest(new SlideMergeResponse
                {
                    Success = false,
                    Message = "Arquivo de apresentação é obrigatório"
                });
            }

            if (!Path.GetExtension(destinationFile.FileName).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new SlideMergeResponse
                {
                    Success = false,
                    Message = "Apenas arquivos .pptx são suportados"
                });
            }

            var result = await _slideMergerService.MergeSlides(destinationFile, request);

            if (!result.Success)
            {
                return BadRequest(result);
            }

            // Retornar o arquivo para download
            if (result.DownloadUrl != null && System.IO.File.Exists(result.DownloadUrl))
            {
                var fileBytes = await System.IO.File.ReadAllBytesAsync(result.DownloadUrl);
                
                // Limpar arquivo temporário após leitura
                System.IO.File.Delete(result.DownloadUrl);
                
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", result.FileName);
            }

            return Ok(result);
        }
    }
}
