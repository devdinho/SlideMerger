using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideMergerAPINew.Models;

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
                if (t.Text.Contains(textoAntigo))
                    t.Text = t.Text.Replace(textoAntigo, textoNovo);
            }
        }

        public static void ApplyPageNumbering(PresentationPart presentationPart)
        {
            var slideIds = presentationPart.Presentation.SlideIdList!.Elements<SlideId>().ToList();

            for (int i = 0; i < slideIds.Count; i++)
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[i].RelationshipId!);
                int pageNumber = i + 1;

                ReplaceTextInSlide(slidePart, "<número>", pageNumber.ToString());
            }
        }

        public static bool SlideWithTextExists(PresentationPart presPart, string searchText)
        {
            return presPart.SlideParts.Any(sp =>
                sp.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                  .Any(t => t.Text.Contains(searchText)));
        }
        public static void ReplaceTextColor(SlidePart slidePart, string textoAlvo, string corHex)
        {
            var textos = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

            foreach (var t in textos)
            {
                if (t.Text.Contains(textoAlvo))
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
            return (long)(17.26  * cmToEmu); // rodapé começa em 17,26cm e tem 1,78cm de altura
        }
        public static bool HasContentOverlappingFooter(SlidePart slidePart, long footerStartY)
        {
            var allElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                .Cast<OpenXmlCompositeElement>()
                .Concat(slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Picture>());

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
                        Console.WriteLine($"Elemento tipo: {element.LocalName}, Y: {yStart / 360000.0}cm, Altura: {height / 360000.0}cm, Y final: {(yStart + height) / 360000.0}cm");
                        return true;
                    }
                }
            }

            return false;
        }

        public static void CopyFooterFromMasterToSlide(SlidePart slidePart)
        {
            var masterPart = slidePart.SlideLayoutPart.SlideMasterPart;
            var footerShapes = masterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                .Where(s =>
                {
                    var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                                .GetFirstChild<PlaceholderShape>();
                    return ph != null && (ph.Type?.Value == PlaceholderValues.Footer ||
                                        ph.Type?.Value == PlaceholderValues.SlideNumber ||
                                        ph.Type?.Value == PlaceholderValues.DateAndTime);
                }).ToList();

            var slideShapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            foreach (var footerShape in footerShapes)
            {
                // Clona o shape
                var clonedShape = (DocumentFormat.OpenXml.Presentation.Shape)footerShape.CloneNode(true);

                // Garante um novo ID único
                uint maxId = slideShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Select(s => s.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value).Max();
                clonedShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = new UInt32Value(maxId + 1);

                // Adiciona ao slide
                slideShapeTree.Append(clonedShape);
            }

            slidePart.Slide.Save();
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

                using var destino = PresentationDocument.Open(outputPath, true);
                using var origem = PresentationDocument.Open(TemplatePath, false);

                var destinoPres = destino.PresentationPart!;
                var origemPres = origem.PresentationPart!;
                var destinoSlides = destinoPres.Presentation.SlideIdList!;
                var origemSlideIds = origemPres.Presentation.SlideIdList!.Elements<SlideId>().ToList();

                int[] primeiros = { 0, 1 };
                int ultimoIdx = origemSlideIds.Count - 1;

                uint NextSlideId() =>
                    destinoSlides.Elements<SlideId>().Max(s => s.Id!.Value) + 1;

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

                Console.WriteLine($"Count: {slidesImgCheck.Count}, Count-: {slidesImgCheck.Count -1}");

                for (int i = 2; i < slidesImgCheck.Count - 2; i++)
                {

                    if (HasContentOverlappingFooter(slidesImgCheck[i], footerLimitY))
                    {
                        CopyFooterFromMasterToSlide(slidesImgCheck[i]);


                        Console.WriteLine($"⚠️ Slide {i + 1} tem conteúdo invadindo a área do rodapé.");
                    }
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

                File.Delete(tempDestinationPath);

                var fileName = $"ApresentacaoFinal_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";

                return new SlideMergeResponse
                {
                    Success = true,
                    Message = "Slides processados com sucesso!",
                    FileName = fileName,
                    DownloadUrl = outputPath
                };
            }
            catch (Exception)
            {
                return new SlideMergeResponse
                {
                    Success = false,
                    Message = $"Erro ao processar slides!"
                };
            }
        }
    }
}
