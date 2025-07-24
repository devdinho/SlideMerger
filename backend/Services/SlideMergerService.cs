using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideMergerAPINew.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;

namespace SlideMergerAPINew.Services
{
    public class SlideMergerService
    {
        private const string TemplatePath = "Templates/Template.pptx";

        public static void ReplaceTextInSlide(SlidePart slidePart, string textoAntigo, string textoNovo)
        {
            var textos = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
            foreach (var t in textos)
            {
                if (t.Text != null && t.Text.Contains(textoAntigo))
                    t.Text = t.Text.Replace(textoAntigo, textoNovo);
            }
        }

        public static void ApplyPageNumbering(PresentationPart presentationPart)
        {
            var slideIds = presentationPart.Presentation.SlideIdList!.Elements<SlideId>().ToList();

            for (int i = 0; i < slideIds.Count; i++)
            {
                var relId = slideIds[i].RelationshipId;
                if (relId == null)
                    continue;

                var slidePart = (SlidePart)presentationPart.GetPartById(relId);
                int pageNumber = i + 1;

                ReplaceTextInSlide(slidePart, "<número>", pageNumber.ToString());
            }
        }

        public static bool SlideWithTextExists(PresentationPart presPart, string searchText)
        {
            return presPart.SlideParts.Any(sp =>
                sp.Slide?.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                    .Any(t => t.Text != null && t.Text.Contains(searchText)) ?? false);
        }

        public static void ReplaceTextColor(SlidePart slidePart, string textoAlvo, string corHex)
        {
            var textos = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

            foreach (var t in textos)
            {
                if (t.Text != null && t.Text.Contains(textoAlvo))
                {
                    var run = t.Parent as Run;
                    if (run != null)
                    {
                        var runProperties = run.RunProperties ??= new RunProperties();

                        runProperties.RemoveAllChildren<SolidFill>();

                        runProperties.AppendChild(new SolidFill(
                            new RgbColorModelHex { Val = corHex.Replace("#", "") }
                        ));
                    }
                }
            }
        }

        public static void ReplaceTextArea(SlidePart slidePart, string textoAlvo, string corHex)
        {
            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();

            foreach (var shape in shapes)
            {
                var texto = shape.TextBody?.InnerText;
                if (!string.IsNullOrEmpty(texto) && texto.Contains(textoAlvo))
                {
                    var spPr = shape.ShapeProperties ??= new DocumentFormat.OpenXml.Presentation.ShapeProperties();

                    spPr.RemoveAllChildren<SolidFill>();

                    spPr.AppendChild(new SolidFill(
                        new RgbColorModelHex { Val = corHex.Replace("#", "") }
                    ));
                }
            }
        }

        public static long GetFooterStartY(PresentationDocument doc)
        {
            const int cmToEmu = 360000;
            return (long)(17.26 * cmToEmu); // rodapé começa em 17,26cm e tem 1,78cm de altura
        }

        // Helper para obter nome do elemento para logging
        private static string GetElementName(OpenXmlCompositeElement element)
        {
            if (element is DocumentFormat.OpenXml.Presentation.Shape shape)
            {
                return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Shape";
            }
            if (element is DocumentFormat.OpenXml.Drawing.Picture picture)
            {
                return picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Picture";
            }
            return element.LocalName;
        }

        public static bool HasContentOverlappingFooter(SlidePart slidePart, long footerStartY)
        {
            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();
            var pictures = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Picture>();
            
            var allElements = shapes.Cast<OpenXmlCompositeElement>().Concat(pictures);

            foreach (var element in allElements)
            {
                var xfrm = element.Descendants<DocumentFormat.OpenXml.Drawing.Transform2D>().FirstOrDefault();
                if (xfrm?.Offset != null && xfrm.Extents != null)
                {
                    long yStart = xfrm.Offset.Y ?? 0;
                    long height = xfrm.Extents.Cy ?? 0;
                    long yEnd = yStart + height;


                    if (yEnd >= footerStartY)
                    {
                        return true; 
                    }
                }
            }
            return false;
        }

        public static void CopyFooterFromMaster(SlidePart slidePart, SlideMasterPart masterPart, long footerYStart)
        {
            var footerShapes = masterPart.SlideMaster.CommonSlideData.ShapeTree
                .OfType<DocumentFormat.OpenXml.Presentation.Shape>()
                .Where(shape =>
                {
                    var transform = shape.ShapeProperties?.Transform2D;
                    if (transform == null) return false;

                    var y = transform.Offset?.Y ?? 0;
                    return y >= footerYStart;
                })
                .ToList();

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            foreach (var footerShape in footerShapes)
            {
                var clonedShape = (DocumentFormat.OpenXml.Presentation.Shape)footerShape.CloneNode(true);

                // (Opcional) Renomear o shape para facilitar depuração
                clonedShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name =
                    new DocumentFormat.OpenXml.StringValue("ClonedFooter_" + Guid.NewGuid().ToString());

                shapeTree.Append(clonedShape); // Adiciona ao final → desenhado por cima
            }

            slidePart.Slide.Save();
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
                var tempDestinationPath = System.IO.Path.GetTempFileName() + ".pptx";
                var outputPath = System.IO.Path.GetTempFileName() + ".pptx";

                using (var stream = new FileStream(tempDestinationPath, FileMode.Create))
                {
                    await destinationFile.CopyToAsync(stream);
                }

                File.Copy(tempDestinationPath, outputPath, true);

                using (var destino = PresentationDocument.Open(outputPath, true))
                using (var origem = PresentationDocument.Open(TemplatePath, false))
                {
                    var destinoPres = destino.PresentationPart!;
                    var origemPres = origem.PresentationPart!;
                    var destinoSlides = destinoPres.Presentation.SlideIdList!;
                    var origemSlideIds = origemPres.Presentation.SlideIdList!.Elements<SlideId>().ToList();

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

                File.Delete(tempDestinationPath);

                string outputNormalizedPath = System.IO.Path.GetTempFileName() + ".pptx";
                await NormalizarComPythonAsync(outputPath, outputNormalizedPath);

                File.Delete(outputPath);

                var fileName = $"{outputNormalizedPath}.pptx";

                return new SlideMergeResponse
                {
                    Success = true,
                    Message = "Slides processados e normalizados com sucesso!",
                    FileName = fileName,
                    DownloadUrl = outputNormalizedPath
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

            using var fs = File.OpenRead(caminhoPptxOriginal);
            using var content = new MultipartFormDataContent();
            using var fileContent = new StreamContent(fs);
            
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            content.Add(fileContent, "file", System.IO.Path.GetFileName(caminhoPptxOriginal));

            var response = await client.PostAsync("http://normalizer:8000/normaliza", content);
            response.EnsureSuccessStatusCode();

            using var ms = await response.Content.ReadAsStreamAsync();
            using var fsOut = File.Create(caminhoPptxNormalizado);
            await ms.CopyToAsync(fsOut);
        }
    }
}