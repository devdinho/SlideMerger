using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideMergerAPINew.Models;
using System;
using System.Collections.Generic;
// Removendo 'using System.IO;' temporariamente para testar se isso resolve a ambiguidade de 'Path' e 'File'
// e explicitando System.IO. na frente dos usos.
// using System.IO; 
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using System.Diagnostics;

// Alias para o namespace Drawing
using D = DocumentFormat.OpenXml.Drawing;

namespace SlideMergerAPINew.Services
{
    public class SlideMergerService
    {
        private const string TemplatePath = "Templates/Template.pptx";

        public static void ReplaceTextInSlide(SlidePart slidePart, string textoAntigo, string textoNovo)
        {
            Console.WriteLine($"Iniciando ReplaceTextInSlide: '{textoAntigo}' por '{textoNovo}'");
            var textos = slidePart.Slide.Descendants<D.Text>();
            foreach (var t in textos)
            {
                if (t.Text != null && t.Text.Contains(textoAntigo))
                    t.Text = t.Text.Replace(textoAntigo, textoNovo);
            }
            Console.WriteLine("Finalizando ReplaceTextInSlide.");
        }

        public static void ApplyPageNumbering(PresentationPart presentationPart)
        {
            Console.WriteLine("Iniciando ApplyPageNumbering.");
            var slideIds = presentationPart.Presentation.SlideIdList!.Elements<SlideId>().ToList();

            for (int i = 0; i < slideIds.Count; i++)
            {
                var relId = slideIds[i].RelationshipId;
                if (string.IsNullOrEmpty(relId))
                    continue;

                var slidePart = (SlidePart)presentationPart.GetPartById(relId);
                int pageNumber = i + 1;
                Console.WriteLine($"  Aplicando numeração de página {pageNumber} ao slide com RelId: {slideIds[i].RelationshipId!}");
                ReplaceTextInSlide(slidePart, "<número>", pageNumber.ToString());
            }
            Console.WriteLine("Finalizando ApplyPageNumbering.");
        }

        public static bool SlideWithTextExists(PresentationPart presPart, string searchText)
        {
            Console.WriteLine($"Verificando se slide com texto '{searchText}' existe.");
            var exists = presPart.SlideParts.Any(sp =>
                sp.Slide?.Descendants<D.Text>()
                    .Any(t => t.Text != null && t.Text.Contains(searchText)) ?? false);
            Console.WriteLine($"  Slide com texto '{searchText}' {(exists ? "encontrado" : "NÃO encontrado")}.");
            return exists;
        }

        public static void ReplaceTextColor(SlidePart slidePart, string textoAlvo, string corHex)
        {
            Console.WriteLine($"Iniciando ReplaceTextColor para '{textoAlvo}' com cor '{corHex}'.");
            var textos = slidePart.Slide.Descendants<D.Text>();

            foreach (var t in textos)
            {
                if (t.Text != null && t.Text.Contains(textoAlvo))
                {
                    var run = t.Parent as D.Run;
                    if (run != null)
                    {
                        var runProperties = run.RunProperties ??= new D.RunProperties();

                        runProperties.RemoveAllChildren<D.SolidFill>();

                        runProperties.AppendChild(new D.SolidFill(
                            new D.RgbColorModelHex { Val = corHex.Replace("#", "") }
                        ));
                        Console.WriteLine($"  Cor alterada para '{textoAlvo}' (Run).");
                    }
                    else
                    {
                        Console.WriteLine($"  AVISO: Pai de texto '{textoAlvo}' não é um Run. Cor não alterada para este texto.");
                    }
                }
            }
            Console.WriteLine("Finalizando ReplaceTextColor.");
        }

        public static void ReplaceTextArea(SlidePart slidePart, string textoAlvo, string corHex)
        {
            Console.WriteLine($"Iniciando ReplaceTextArea para '{textoAlvo}' com cor de área '{corHex}'.");
            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();

            foreach (var shape in shapes)
            {
                var texto = shape.TextBody?.InnerText;
                if (!string.IsNullOrEmpty(texto) && texto.Contains(textoAlvo))
                {
                    var spPr = shape.ShapeProperties ??= new DocumentFormat.OpenXml.Presentation.ShapeProperties();

                    spPr.RemoveAllChildren<D.SolidFill>();

                    spPr.AppendChild(new D.SolidFill(
                        new D.RgbColorModelHex { Val = corHex.Replace("#", "") }
                    ));
                    Console.WriteLine($"  Cor da área alterada para shape com texto '{textoAlvo}'.");
                }
            }
            Console.WriteLine("Finalizando ReplaceTextArea.");
        }

        public static long GetFooterStartY(PresentationDocument doc)
        {
            const int cmToEmu = 360000;
            return (long)(17.26 * cmToEmu); // rodapé começa em 17,26cm e tem 1,78cm de altura
        }

        private static long? GetYPosition(OpenXmlElement element)
        {
            var xfrm = element.Descendants<D.Transform2D>().FirstOrDefault();
            return xfrm?.Offset?.Y?.Value;
        }

        private static string GetElementName(OpenXmlElement element)
        {
            if (element is DocumentFormat.OpenXml.Presentation.Shape shape)
            {
                return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Shape";
            }
            if (element is D.Picture picture)
            {
                return picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Picture";
            }
            if (element is DocumentFormat.OpenXml.Presentation.GroupShape groupShape)
            {
                return groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed GroupShape";
            }
            return element.LocalName;
        }

        public static bool HasContentOverlappingFooter(SlidePart slidePart, long footerStartY)
        {
            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();
            var pictures = slidePart.Slide.Descendants<D.Picture>();
            
            var allElements = shapes.Cast<OpenXmlElement>().Concat(pictures.Cast<OpenXmlElement>());

            foreach (var element in allElements)
            {
                var xfrm = element.Descendants<D.Transform2D>().FirstOrDefault();
                if (xfrm?.Offset != null && xfrm.Extents != null)
                {
                    long yStart = xfrm.Offset.Y?.Value ?? 0;
                    long height = xfrm.Extents.Cy?.Value ?? 0;
                    long yEnd = yStart + height;

                    if (yEnd >= footerStartY)
                    {
                        Console.WriteLine($"[DEBUG] Sobreposição detectada no slide: {slidePart.Uri.OriginalString}. Elemento sobreposto: {GetElementName(element)}, Y inicial: {yStart/360000.0:F3}cm, Altura: {height/360000.0:F3}cm, Y final: {yEnd/360000.0:F3}cm");
                        return true; 
                    }
                }
            }
            Console.WriteLine($"[DEBUG] Nenhuma sobreposição detectada no slide: {slidePart.Uri.OriginalString}.");
            return false;
        }

        private static IEnumerable<OpenXmlElement> GetShapesAndPicturesFromElement(OpenXmlElement element)
        {
            if (element == null) yield break;

            if (element is DocumentFormat.OpenXml.Presentation.Shape || element is D.Picture)
            {
                Console.WriteLine($"[DEBUG] Elemento final avaliando para cópia: '{element.LocalName}' (Tipo XML: '{element.GetType().Name}'), Y: {GetYPosition(element).ToCm():F3}cm (EMUs raw: {GetYPosition(element).GetValueOrDefault()})");
                yield return element;
            }

            if (element is OpenXmlCompositeElement compositeElement)
            {
                foreach (var child in compositeElement.Elements())
                {
                    foreach (var s in GetShapesAndPicturesFromElement(child))
                    {
                        yield return s;
                    }
                }
            }
        }

        public static void CopyFooterFromMaster(SlidePart slidePart, SlideMasterPart masterPart, long footerYStart)
        {
            Console.WriteLine($"[DEBUG] Buscando shapes de rodapé no Slide Master. Filtrando a partir de {footerYStart/360000.0:F3}cm.");

            Console.WriteLine("[DEBUG] Conteúdo completo do Slide Master ShapeTree:");
            var allMasterRootElements = masterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlElement>();
            foreach (var element in allMasterRootElements)
            {
                var yPos = GetYPosition(element);
                long y = yPos.GetValueOrDefault(-1); 
                string type = element.LocalName;
                string name = GetElementName(element);
                Console.WriteLine($"  - Elemento: '{name}' (Tipo XML: '{type}'), Y: {y/360000.0:F3}cm (EMUs raw: {y})");
            }
            Console.WriteLine("[DEBUG] Fim do Conteúdo completo do Slide Master ShapeTree.");

            var allShapesAndPicturesInMaster = GetShapesAndPicturesFromElement(masterPart.SlideMaster.CommonSlideData.ShapeTree);

            var footerElementsToCopy = new List<OpenXmlElement>();

            foreach(var element in allShapesAndPicturesInMaster)
            {
                var yPos = GetYPosition(element);
                if (yPos.HasValue)
                {
                    long y = yPos.Value;
                    Console.WriteLine($"[DEBUG] Elemento final avaliando para cópia: '{GetElementName(element)}' (Tipo XML: '{element.LocalName}'), Y: {y/360000.0:F3}cm. (Será copiado se Y >= {footerYStart/360000.0:F3}cm)");
                    
                    if (y >= footerYStart)
                    {
                        footerElementsToCopy.Add(element);
                    }
                } else {
                    Console.WriteLine($"[DEBUG] Elemento final: '{GetElementName(element)}' (Tipo XML: '{element.LocalName}') não tem Transform2D ou Offset e será ignorado pelo filtro.");
                }
            }

            if (!footerElementsToCopy.Any())
            {
                Console.WriteLine($"[DEBUG] Nenhuma forma de rodapé encontrada no Slide Master para copiar com Y >= {footerYStart/360000.0:F3}cm.");
            }

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            foreach (var footerElement in footerElementsToCopy)
            {
                var clonedElement = (OpenXmlElement)footerElement.CloneNode(true);

                if (clonedElement is DocumentFormat.OpenXml.Presentation.Shape clonedShape)
                {
                    clonedShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name =
                        new DocumentFormat.OpenXml.StringValue("ClonedFooterShape_" + Guid.NewGuid().ToString());
                }
                else if (clonedElement is D.Picture clonedPicture)
                {
                    clonedPicture.NonVisualPictureProperties.NonVisualDrawingProperties.Name =
                        new DocumentFormat.OpenXml.StringValue("ClonedFooterPicture_" + Guid.NewGuid().ToString());
                }

                shapeTree.Append(clonedElement);
                Console.WriteLine($"[DEBUG] Rodapé '{GetElementName(clonedElement)}' copiado para o slide {slidePart.Uri.OriginalString}.");
            }

            slidePart.Slide.Save();
            Console.WriteLine($"[DEBUG] Slide {slidePart.Uri.OriginalString} salvo após tentativa de cópia do rodapé.");
        }


        public static void AddFooterOverlayToSlide(SlidePart slidePart, SlideMasterPart masterPart, long footerLimitY)
        {
            if (HasContentOverlappingFooter(slidePart, footerLimitY))
            {
                CopyFooterFromMaster(slidePart, masterPart, footerLimitY - (long)(1.78 * 360000));
            }
        }

        public async Task<SlideMergeResponse> MergeSlides(IFormFile destinationFile, SlideMergeRequest request)
        {
            try
            {
                // **Explicitando System.IO.Path e System.IO.File**
                string tempDestinationPath = System.IO.Path.GetTempFileName() + ".pptx";
                string outputPath = System.IO.Path.GetTempFileName() + ".pptx";

                using (var stream = new System.IO.FileStream(tempDestinationPath, System.IO.FileMode.Create))
                {
                    await destinationFile.CopyToAsync(stream);
                }

                System.IO.File.Copy(tempDestinationPath, outputPath, true);

                using (var destino = PresentationDocument.Open(outputPath, true))
                using (var origem = PresentationDocument.Open(TemplatePath, false))
                {
                    var destinoPres = destino.PresentationPart!;
                    var origemPres = origem.PresentationPart!;
                    var destinoSlides = destinoPres.Presentation!.SlideIdList!;
                    var origemSlideIds = origemPres.Presentation!.SlideIdList!.Elements<SlideId>().ToList();

                    int[] primeiros = { 0, 1 };
                    int ultimoIdx = origemSlideIds.Count - 1;

                    uint NextSlideId() =>
                        destinoSlides.Elements<SlideId>().Any() ? 
                        destinoSlides.Elements<SlideId>().Max(s => s.Id!.Value) + 1 : 256U;

                    foreach (int i in primeiros)
                    {
                        var origemSlidePart = (SlidePart)origemPres.GetPartById(origemSlideIds[i].RelationshipId!);

                        string marcadorVerificacao = i == 0 ? "Prof" : "Lei nº 9610/98";

                        if (SlideWithTextExists(destinoPres, marcadorVerificacao))
                        {
                            continue;
                        }

                        var novoSlidePart = destinoPres.AddPart(origemSlidePart);
                        uint novoId = NextSlideId();

                        ReplaceTextColor(novoSlidePart, "NOMEMBA", request.Theme);
                        ReplaceTextInSlide(novoSlidePart, "NOMEMBA", request.Mba.ToUpper());
                        ReplaceTextInSlide(novoSlidePart, "Título da aula/disciplina", request.TituloAula);
                        ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");

                        ReplaceTextArea(novoSlidePart, "Lei nº 9610/98", request.Theme);

                        destinoSlides.InsertAt(new SlideId
                        {
                            Id = novoId,
                            RelationshipId = destinoPres.GetIdOfPart(novoSlidePart)
                        }, i);
                    }

                    if (!SlideWithTextExists(destinoPres, "linkedin.com/in/"))
                    {
                        var origemSlidePart = (SlidePart)origemPres.GetPartById(origemSlideIds[ultimoIdx].RelationshipId!);
                        var novoSlidePart = destinoPres.AddPart(origemSlidePart);
                        ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");
                        ReplaceTextColor(novoSlidePart, "linkedin.perfil.com", request.Theme);
                        ReplaceTextInSlide(novoSlidePart, "linkedin.perfil.com", request.LinkedinPerfil);

                        destinoSlides.Append(new SlideId
                        {
                            Id = NextSlideId(),
                            RelationshipId = destinoPres.GetIdOfPart(novoSlidePart)
                        });
                    }

                    var thirdSlide = (SlidePart)origemPres.GetPartById(origemSlideIds[2].RelationshipId!);
                    var layoutSrc = thirdSlide.SlideLayoutPart!;
                    var masterSrc = layoutSrc.SlideMasterPart!;
                    var layoutName = layoutSrc.SlideLayout.Type?.Value.ToString() ?? "default";

                    var existingMaster = destinoPres.SlideMasterParts
                        .FirstOrDefault(mp => mp.SlideMaster.Descendants<SlideLayout>()
                            .Any(l => (l.Type?.Value.ToString() ?? "default") == layoutName));

                    SlideMasterPart masterDest = existingMaster ?? destinoPres.AddPart(masterSrc);

                    if (existingMaster == null)
                    {
                        foreach (var lay in masterSrc.SlideLayoutParts)
                            masterDest.AddPart(lay);

                        var masterIdList = destinoPres.Presentation.SlideMasterIdList
                                         ?? destinoPres.Presentation.AppendChild(new SlideMasterIdList());

                        uint nextMasterId = masterIdList.Elements<SlideMasterId>().Any()
                            ? masterIdList.Elements<SlideMasterId>().Max(m => m.Id!.Value) + 1
                            : 1;

                        masterIdList.Append(new SlideMasterId
                        {
                            Id = nextMasterId,
                            RelationshipId = destinoPres.GetIdOfPart(masterDest)
                        });
                    }

                    var layoutDest = masterDest.SlideLayoutParts
                        .First(l => (l.SlideLayout.Type?.Value.ToString() ?? "default") == layoutName);

                    var slideIds = destinoSlides.Elements<SlideId>().ToList();
                    var slidesImgCheck = destino.PresentationPart!.SlideParts.ToList();

                    var footerLimitY = GetFooterStartY(destino);


                    for (int i = 2; i < slidesImgCheck.Count - 2; i++)
                    {
                        AddFooterOverlayToSlide(slidesImgCheck[i], masterSrc, footerLimitY);
                    }

                    for (int i = 2; i < slideIds.Count - 1; i++)
                    {
                        var slidePart = (SlidePart)destinoPres.GetPartById(slideIds[i].RelationshipId!);

                        if (slidePart.SlideLayoutPart != null)
                            slidePart.DeletePart(slidePart.SlideLayoutPart);

                        slidePart.AddPart(layoutDest);
                        slidePart.Slide.Save();
                    }

                    ApplyPageNumbering(destinoPres);
                    destinoPres.Presentation.Save();
                }

                System.IO.File.Delete(tempDestinationPath);

                string outputNormalizedPath = System.IO.Path.GetTempFileName() + ".pptx";
                await NormalizarComPythonAsync(outputPath, outputNormalizedPath);

                System.IO.File.Delete(outputPath);

                var fileName = System.IO.Path.GetFileName(outputNormalizedPath);
                var downloadUrl = outputNormalizedPath;


                return new SlideMergeResponse
                {
                    Success = true,
                    Message = "Slides processados e normalizados com sucesso!",
                    FileName = fileName,
                    DownloadUrl = downloadUrl
                };
            }
            catch (Exception ex)
            {
                return new SlideMergeResponse
                {
                    Success = false,
                    Message = $"Erro ao processar slides! Detalhes: {ex.Message}"
                };
            }
        }

        private static async Task NormalizarComPythonAsync(string caminhoPptxOriginal, string caminhoPptxNormalizado)
        {
            using var client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(5);

            using var fs = System.IO.File.OpenRead(caminhoPptxOriginal);
            using var content = new MultipartFormDataContent();
            using var fileContent = new StreamContent(fs);
            
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            content.Add(fileContent, "file", System.IO.Path.GetFileName(caminhoPptxOriginal));

            var response = await client.PostAsync("http://normalizer:8000/normaliza", content);
            response.EnsureSuccessStatusCode();

            using var ms = await response.Content.ReadAsStreamAsync();
            using var fsOut = System.IO.File.Create(caminhoPptxNormalizado);
            await ms.CopyToAsync(fsOut);
        }
    }

    public static class OpenXmlHelperExtensions
    {
        public static double ToCm(this long? emuValue)
        {
            if (emuValue.HasValue)
            {
                return emuValue.Value / 36000.0; // 1 cm = 36000 EMUs
            }
            return 0.0;
        }
    }
}