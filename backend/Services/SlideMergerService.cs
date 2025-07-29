using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideMergerAPINew.Models;
using System;
using System.Collections.Generic;
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
            var textos = slidePart.Slide.Descendants<D.Text>();
            foreach (var t in textos)
            {
                if (t.Text?.Contains(textoAntigo) == true)
                    t.Text = t.Text!.Replace(textoAntigo, textoNovo);
            }
        }

        public static void ApplyPageNumbering(PresentationPart presentationPart)
        {
            var slideIds = presentationPart.Presentation.SlideIdList!.Elements<SlideId>().ToList();

            for (int i = 0; i < slideIds.Count; i++)
            {
                var relId = slideIds[i].RelationshipId;
                if (string.IsNullOrEmpty(relId))
                    continue;

                var slidePart = (SlidePart)presentationPart.GetPartById(relId!);
                int pageNumber = i + 1;
                ReplaceTextInSlide(slidePart, "<número>", pageNumber.ToString());
            }
        }

        public static bool SlideWithTextExists(PresentationPart presPart, string searchText)
        {
            var exists = presPart.SlideParts.Any(sp =>
                sp.Slide?.Descendants<D.Text>()
                    .Any(t => t.Text?.Contains(searchText) == true) ?? false);
            return exists;
        }

        public static bool SlideWithTextExists(SlidePart slidePart, string searchText)
        {
            return slidePart.Slide?.Descendants<D.Text>()
                .Any(t => t.Text?.Contains(searchText) == true) ?? false;
        }

        public static void ReplaceTextColor(SlidePart slidePart, string textoAlvo, string corHex)
        {
            var textos = slidePart.Slide.Descendants<D.Text>();

            foreach (var t in textos)
            {
                if (t.Text?.Contains(textoAlvo) == true)
                {
                    var run = t.Parent as D.Run;
                    if (run != null)
                    {
                        var runProperties = run.RunProperties ??= new D.RunProperties();

                        runProperties.RemoveAllChildren<D.SolidFill>();

                        runProperties.AppendChild(new D.SolidFill(
                            new D.RgbColorModelHex { Val = corHex.Replace("#", "") }
                        ));
                    }
                    else
                    {
                        Console.WriteLine($"AVISO: Pai de texto '{textoAlvo}' não é um Run. Cor não alterada para este texto.");
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

                    spPr.RemoveAllChildren<D.SolidFill>();

                    spPr.AppendChild(new D.SolidFill(
                        new D.RgbColorModelHex { Val = corHex.Replace("#", "") }
                    ));
                }
            }
        }

        public static long GetFooterStartY(PresentationDocument doc)
        {
            const int cmToEmu = 360000;
            return (long)(17.26 * cmToEmu); // rodapé começa em 17,26cm e tem 1,78cm de altura
        }

        private static long? GetYPosition(OpenXmlElement element)
        {
            if (element is DocumentFormat.OpenXml.Presentation.GroupShape gShape)
            {
                var groupOffset = gShape.Descendants<D.TransformGroup>().FirstOrDefault()?.Offset;
                return groupOffset?.Y?.Value;
            }

            var xfrm = element.Descendants<D.Transform2D>().FirstOrDefault();
            return xfrm?.Offset?.Y?.Value;
        }

        private static string GetElementName(OpenXmlElement element)
        {
            if (element is DocumentFormat.OpenXml.Presentation.Shape shape)
            {
                return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Shape";
            }
            else if (element is D.Picture picture)
            {
                return picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed Picture";
            }
            else if (element is DocumentFormat.OpenXml.Presentation.GroupShape groupShape)
            {
                return groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed GroupShape";
            }
            else if (element is DocumentFormat.OpenXml.Presentation.GraphicFrame graphicFrame)
            {
                var graphicData = graphicFrame.Graphic?.GraphicData;
                if (graphicData != null && graphicData.Uri?.Value == "http://schemas.microsoft.com/office/drawing/2010/svg")
                {
                    return graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed SVG GraphicFrame";
                }
                return graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Unnamed GraphicFrame";
            }
            return element.LocalName;
        }

        public static bool HasContentOverlappingFooter(SlidePart slidePart, long footerStartY)
        {
            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();
            var pictures = slidePart.Slide.Descendants<D.Picture>();
            var graphicFrames = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>();
            
            var allElements = shapes.Cast<OpenXmlElement>()
                                    .Concat(pictures.Cast<OpenXmlElement>())
                                    .Concat(graphicFrames.Cast<OpenXmlElement>());

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

        public static void CopyFooterFromMaster(SlidePart slidePart, SlideMasterPart masterPart, PresentationDocument destinationDocument, long footerYStart)
        {
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            var footerElementsToCopy = new List<OpenXmlElement>();

            var allMasterVisualDescendants = masterPart.SlideMaster!.CommonSlideData!.ShapeTree!.Descendants<OpenXmlElement>()
                .Where(e => e is DocumentFormat.OpenXml.Presentation.Shape || 
                            e is D.Picture || 
                            e is DocumentFormat.OpenXml.Presentation.GroupShape || 
                            e is DocumentFormat.OpenXml.Presentation.GraphicFrame)
                .ToList();

            var addedElementsTracker = new HashSet<OpenXmlElement>();

            Console.WriteLine("[DEBUG] Elementos visuais detectados no Slide Master (todos os tipos):");
            foreach (var dbgElement in allMasterVisualDescendants)
            {
                long? y = GetYPosition(dbgElement);
                string elementName = GetElementName(dbgElement); 
                Console.WriteLine($"  - Elemento: {elementName} (Tipo: {dbgElement.LocalName}), Y: {(y.HasValue ? y.Value / 360000.0 : -1):F3}cm");
            }
            Console.WriteLine("[DEBUG] Fim da lista de elementos visuais do Slide Master.");

            foreach (var element in allMasterVisualDescendants)
            {
                bool shouldCopy = true; 

                if (addedElementsTracker.Contains(element))
                {
                    continue; 
                }
                
                Console.WriteLine($"[DEBUG] Considerando para cópia: {GetElementName(element)} (Tipo: {element.LocalName}).");

                if (shouldCopy && !addedElementsTracker.Contains(element))
                {
                    footerElementsToCopy.Add(element);
                    addedElementsTracker.Add(element);
                    
                    if (element is DocumentFormat.OpenXml.Presentation.GroupShape groupBeingAdded)
                    {
                        foreach (var descendant in groupBeingAdded.Descendants<OpenXmlElement>())
                        {
                            addedElementsTracker.Add(descendant);
                        }
                        Console.WriteLine($"[DEBUG] Grupo '{GetElementName(element)}' adicionado para cópia. Seus filhos serão ignorados.");
                    }
                    else
                    {
                        Console.WriteLine($"[DEBUG] Elemento '{GetElementName(element)}' adicionado para cópia.");
                    }
                }
            }

            if (!footerElementsToCopy.Any())
            {
                Console.WriteLine($"[DEBUG] Nenhum elemento foi selecionado para cópia do master.");
                return;
            }

            foreach (var footerElement in footerElementsToCopy)
            {
                var clonedElement = (OpenXmlElement)footerElement.CloneNode(true);

                // --- Atribui nomes únicos e lida com partes de imagem (SVG, PNG, JPG, etc.) ---
                if (clonedElement is DocumentFormat.OpenXml.Presentation.Shape shape)
                {
                    shape.NonVisualShapeProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties();
                    shape.NonVisualShapeProperties.NonVisualDrawingProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterShape_" + Guid.NewGuid().ToString());
                }
                else if (clonedElement is D.Picture pic) 
                {
                    string uniqueName = "ClonedFooterPicture_" + Guid.NewGuid().ToString();
                    pic.NonVisualPictureProperties ??= new D.NonVisualPictureProperties();
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties ??= new D.NonVisualDrawingProperties();
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name = new StringValue(uniqueName);

                    var originalPic = footerElement as D.Picture;
                    if (originalPic != null && originalPic.BlipFill?.Blip?.Embed?.Value != null)
                    {
                        string originalImageRelId = originalPic.BlipFill.Blip.Embed.Value;
                        ImagePart originalImagePart = (ImagePart)masterPart.GetPartById(originalImageRelId);

                        if (originalImagePart != null)
                        {
                            // === CORREÇÃO: Usar AddNewPart<ImagePart>() para adicionar a imagem ===
                            ImagePart newImagePart = destinationDocument.PresentationPart.AddNewPart<ImagePart>(originalImagePart.ContentType);
                            using (var stream = originalImagePart.GetStream())
                            {
                                stream.CopyTo(newImagePart.GetStream());
                            }
                            // Obter o ID da relação da PresentationPart, que é quem gerencia as relações de imagem
                            pic.BlipFill.Blip.Embed = destinationDocument.PresentationPart.GetIdOfPart(newImagePart);
                            Console.WriteLine($"[DEBUG] Imagem '{GetElementName(originalPic)}' copiada e RelationshipId atualizado para '{destinationDocument.PresentationPart.GetIdOfPart(newImagePart)}'.");
                        }
                    }
                }
                else if (clonedElement is DocumentFormat.OpenXml.Presentation.GroupShape group)
                {
                    group.NonVisualGroupShapeProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties();
                    group.NonVisualGroupShapeProperties.NonVisualDrawingProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    group.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterGroup_" + Guid.NewGuid().ToString());

                    foreach (var descendantBlip in group.Descendants<D.Blip>())
                    {
                        if (descendantBlip?.Embed?.Value != null)
                        {
                            string originalImageRelId = descendantBlip.Embed.Value;
                            var originalImagePart = masterPart.GetPartById(originalImageRelId) as ImagePart;

                            if (originalImagePart != null)
                            {
                                // === CORREÇÃO: Usar AddNewPart<ImagePart>() para adicionar a imagem ===
                                ImagePart newImagePart = destinationDocument.PresentationPart.AddNewPart<ImagePart>(originalImagePart.ContentType);
                                using (var stream = originalImagePart.GetStream())
                                {
                                    stream.CopyTo(newImagePart.GetStream());
                                }
                                descendantBlip.Embed = destinationDocument.PresentationPart.GetIdOfPart(newImagePart);
                                Console.WriteLine($"[DEBUG] Imagem aninhada em grupo com ID '{originalImageRelId}' copiada e RelationshipId atualizado.");
                            }
                        }
                    }
                }
                else if (clonedElement is DocumentFormat.OpenXml.Presentation.GraphicFrame graphicFrame)
                {
                    string uniqueName = "ClonedFooterGraphicFrame_" + Guid.NewGuid().ToString();
                    graphicFrame.NonVisualGraphicFrameProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameProperties();
                    graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = new StringValue(uniqueName);

                    var originalGraphicFrame = footerElement as DocumentFormat.OpenXml.Presentation.GraphicFrame;
                    if (originalGraphicFrame != null)
                    {
                        var blip = clonedElement.Descendants<D.Blip>().FirstOrDefault();
                        if (blip?.Embed?.Value != null)
                        {
                            string originalImageRelId = blip.Embed.Value;
                            ImagePart originalImagePart = (ImagePart)masterPart.GetPartById(originalImageRelId);

                            if (originalImagePart != null)
                            {
                                // === CORREÇÃO: Usar AddNewPart<ImagePart>() para adicionar a imagem ===
                                ImagePart newImagePart = destinationDocument.PresentationPart.AddNewPart<ImagePart>(originalImagePart.ContentType);
                                using (var stream = originalImagePart.GetStream())
                                {
                                    stream.CopyTo(newImagePart.GetStream());
                                }
                                blip.Embed = destinationDocument.PresentationPart.GetIdOfPart(newImagePart);
                                Console.WriteLine($"[DEBUG] SVG/Imagem '{GetElementName(originalGraphicFrame)}' copiado e RelationshipId atualizado para '{destinationDocument.PresentationPart.GetIdOfPart(newImagePart)}'.");
                            }
                        }
                    }
                }

                shapeTree.Append(clonedElement);
                Console.WriteLine($"[DEBUG] Elemento '{GetElementName(clonedElement)}' copiado para o slide {slidePart.Uri.OriginalString}.");
            }

            slidePart.Slide!.Save();
            Console.WriteLine($"[DEBUG] Slide {slidePart.Uri.OriginalString} salvo após tentativa de cópia do rodapé.");
        }

        public static void AddFooterOverlayToSlide(SlidePart slidePart, SlideMasterPart masterPart, long footerLimitY)
        {
            Console.WriteLine($"[DEBUG] Chamando CopyFooterFromMaster incondicionalmente para o slide {slidePart.Uri.OriginalString}.");
            
            var destinationDocument = slidePart.OpenXmlPackage as PresentationDocument; 
            
            if (destinationDocument?.PresentationPart == null) 
            {
                Console.WriteLine("[ERROR] PresentationPart de destino é nula. Não é possível copiar o rodapé.");
                return;
            }

            CopyFooterFromMaster(slidePart, masterPart, destinationDocument, footerLimitY);
        }

        public async Task<SlideMergeResponse> MergeSlides(IFormFile destinationFile, SlideMergeRequest request)
        {
            try
            {
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
                        .FirstOrDefault(mp => mp.SlideMaster?.Descendants<SlideLayout>()
                            .Any(l => (l.Type?.Value.ToString() ?? "default") == layoutName) ?? false);

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

                    var currentSlideIdsInDest = destinoSlides.Elements<SlideId>().ToList();
                    
                    var slidesForFooterCheck = destino.PresentationPart!.SlideParts.ToList();

                    var footerLimitY = GetFooterStartY(destino);

                    for (int i = 0; i < currentSlideIdsInDest.Count; i++)
                    {
                        var slidePart = (SlidePart)destinoPres.GetPartById(currentSlideIdsInDest[i].RelationshipId!);

                        if (i < 2 || i == currentSlideIdsInDest.Count - 1)
                        {
                             Console.WriteLine($"[DEBUG] Ignorando aplicação de layout ou rodapé para slide index {i}.");
                             continue;
                        }

                        if (slidePart.SlideLayoutPart != null)
                            slidePart.DeletePart(slidePart.SlideLayoutPart);
                        slidePart.AddPart(layoutDest);
                        slidePart.Slide!.Save();

                        if (!SlideWithTextExists(slidePart, "Prof") && !SlideWithTextExists(slidePart, "Lei nº 9610/98"))
                        {
                            AddFooterOverlayToSlide(slidePart, masterSrc, footerLimitY);
                        }
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