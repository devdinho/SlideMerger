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
            if (slidePart.Slide == null) return; // Add null check for CS8602
            var textos = slidePart.Slide.Descendants<D.Text>();
            foreach (var t in textos)
            {
                if (t.Text?.Contains(textoAntigo) == true)
                    t.Text = t.Text!.Replace(textoAntigo, textoNovo);
            }
        }

        public static void ApplyPageNumbering(PresentationPart presentationPart)
        {
            var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();

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
            if (slidePart.Slide == null) return; // Add null check for CS8602
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
            var shapes = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();

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
            return (long)(17.0 * cmToEmu);
        }

        private static long? GetYPosition(OpenXmlElement element)
        {
            var transform = element.Descendants<D.Transform2D>().FirstOrDefault();
            if (transform != null) return transform.Offset?.Y?.Value;

            var groupTransform = element.Descendants<D.TransformGroup>().FirstOrDefault();
            if (groupTransform != null) return groupTransform.Offset?.Y?.Value;
            
            return null;
        }
        
        private static long? GetElementHeight(OpenXmlElement element)
        {
            var transform = element.Descendants<D.Transform2D>().FirstOrDefault();
            if (transform != null) return transform.Extents?.Cy?.Value;

            var groupTransform = element.Descendants<D.TransformGroup>().FirstOrDefault();
            if (groupTransform != null) return groupTransform.Extents?.Cy?.Value;

            return null;
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
                return groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Grupo Sem Nome";
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
            var nameProperty = element.Descendants<D.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value;
            if (nameProperty != null) return nameProperty;

            return element.LocalName;
        }

        public static bool HasContentOverlappingFooter(SlidePart slidePart, long footerStartY)
        {
            var shapes = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();
            var pictures = slidePart.Slide!.Descendants<D.Picture>();
            var graphicFrames = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>();
            var groupShapes = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.GroupShape>();
            
            var allElements = shapes.Cast<OpenXmlElement>()
                                        .Concat(pictures.Cast<OpenXmlElement>())
                                        .Concat(graphicFrames.Cast<OpenXmlElement>())
                                        .Concat(groupShapes.Cast<OpenXmlElement>());

            foreach (var element in allElements)
            {
                long? yStart = GetYPosition(element);
                long? height = GetElementHeight(element);

                if (yStart.HasValue && height.HasValue)
                {
                    long yEnd = yStart.Value + height.Value;

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

        public static void CopyFooterFromMaster(SlidePart slidePart, SlideMasterPart masterPart, long footerYStart) // Removed PresentationDocument parameter as it was not needed
        {
            var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            var footerElementsToCopy = new List<OpenXmlElement>();

            var allMasterVisualDescendants = masterPart.SlideMaster!.Descendants<OpenXmlElement>()
                .Where(e => (e is DocumentFormat.OpenXml.Presentation.Shape ||
                              e is D.Picture ||
                              e is DocumentFormat.OpenXml.Presentation.GraphicFrame ||
                              e is DocumentFormat.OpenXml.Presentation.GroupShape) &&
                             GetYPosition(e).HasValue && GetElementHeight(e).HasValue)
                .ToList();

            var addedElementsTracker = new HashSet<OpenXmlElement>();

            Console.WriteLine("[DEBUG] Elementos visuais detectados no Slide Master (todos os tipos):");
            foreach (var dbgElement in allMasterVisualDescendants)
            {
                long? y = GetYPosition(dbgElement);
                long? height = GetElementHeight(dbgElement);
                string elementName = GetElementName(dbgElement);
                string innerText = (dbgElement as DocumentFormat.OpenXml.Presentation.Shape)?.TextBody?.InnerText ?? "";
                Console.WriteLine($"  - Elemento: {elementName} (Tipo: {dbgElement.LocalName}), Y: {(y.HasValue ? y.Value / 360000.0 : -1):F3}cm, Altura: {(height.HasValue ? height.Value / 360000.0 : -1):F3}cm, Texto: '{innerText.Replace("\n", " ").Replace("\r", "")}'");
            }
            Console.WriteLine("[DEBUG] Fim da lista de elementos visuais do Slide Master.");

            var desiredFooterElementIdentifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Retângulo 11", 
                "CaixaDeTexto 13", 
                "PlaceHolder 3",   
                "Gráfico 15",  
                "Conector reto 14" 
            };

            foreach (var element in allMasterVisualDescendants)
            {
                if (addedElementsTracker.Contains(element))
                {
                    continue;
                }

                long? yStart = GetYPosition(element);
                long? height = GetElementHeight(element);
                string elementName = GetElementName(element);
                string innerText = (element as DocumentFormat.OpenXml.Presentation.Shape)?.TextBody?.InnerText ?? "";
                innerText = innerText.Replace("\n", " ").Replace("\r", ""); 

                bool isFooterCandidateByPosition = false;
                if (yStart.HasValue && height.HasValue)
                {
                    long yEnd = yStart.Value + height.Value;
                    if (yStart.Value >= footerYStart || yEnd >= footerYStart)
                    {
                        isFooterCandidateByPosition = true;
                    }
                }
                
                if (isFooterCandidateByPosition)
                {
                    if (element is DocumentFormat.OpenXml.Presentation.GroupShape groupBeingAdded)
                    {
                        footerElementsToCopy.Add(element);
                        addedElementsTracker.Add(element); 

                        foreach (var descendant in groupBeingAdded.Descendants<OpenXmlElement>())
                        {
                            addedElementsTracker.Add(descendant);
                        }
                        Console.WriteLine($"[DEBUG] Grupo '{elementName}' adicionado para cópia (Y inicial: {(yStart.HasValue ? yStart.Value / 360000.0 : -1):F3}cm). Seus filhos serão ignorados individualmente na detecção principal, but will be copied as part of the group.");
                    }
                    else
                    {
                        footerElementsToCopy.Add(element);
                        addedElementsTracker.Add(element);
                        Console.WriteLine($"[DEBUG] Elemento '{elementName}' adicionado para cópia (Y inicial: {(yStart.HasValue ? yStart.Value / 360000.0 : -1):F3}cm, Texto: '{innerText}').");
                    }
                }
                else
                {
                    Console.WriteLine($"[DEBUG] Elemento '{elementName}' (Y inicial: {(yStart.HasValue ? yStart.Value / 360000.0 : -1):F3}cm, Altura: {(height.HasValue ? height.Value / 360000.0 : -1):F3}cm, Texto: '{innerText}') ignorado, pois não está na área do rodapé (início do rodapé: {footerYStart/360000.0:F3}cm).");
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

                if (clonedElement is DocumentFormat.OpenXml.Presentation.Shape shape)
                {
                    shape.NonVisualShapeProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties();
                    if (shape.NonVisualShapeProperties.NonVisualDrawingProperties == null)
                    {
                        shape.NonVisualShapeProperties.NonVisualDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    }
                    shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterShape_" + Guid.NewGuid().ToString());
                }
                else if (clonedElement is D.Picture pic)
                {
                    string uniqueName = "ClonedFooterPicture_" + Guid.NewGuid().ToString();
                    pic.NonVisualPictureProperties ??= new D.NonVisualPictureProperties();
                    if (pic.NonVisualPictureProperties.NonVisualDrawingProperties == null)
                    {
                        pic.NonVisualPictureProperties.NonVisualDrawingProperties = new D.NonVisualDrawingProperties(); 
                    }
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name = new StringValue(uniqueName);

                    var originalPic = footerElement as D.Picture;
                    if (originalPic != null && originalPic.BlipFill?.Blip?.Embed?.Value != null)
                    {
                        string originalImageRelId = originalPic.BlipFill.Blip.Embed.Value;
                        ImagePart originalImagePart = (ImagePart)masterPart.GetPartById(originalImageRelId);

                        if (originalImagePart != null)
                        {
                            Console.WriteLine($"[DEBUG] CopyFooterFromMaster: Adding new ImagePart to SlidePart with ContentType: '{originalImagePart.ContentType}' and explicit ID.");
                            string newRelId = "rId" + Guid.NewGuid().ToString("N");
                            ImagePart newImagePart = slidePart.AddNewPart<ImagePart>(originalImagePart.ContentType, newRelId); 
                            using (var stream = originalImagePart.GetStream())
                            {
                                stream.CopyTo(newImagePart.GetStream());
                            }
                            pic.BlipFill.Blip.Embed = newRelId; // Use the explicitly generated ID
                            Console.WriteLine($"[DEBUG] Imagem '{GetElementName(originalPic)}' copiada e RelationshipId atualizado para '{newRelId}'.");
                        }
                    }
                }
                else if (clonedElement is DocumentFormat.OpenXml.Presentation.GroupShape group)
                {
                    group.NonVisualGroupShapeProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties();
                    if (group.NonVisualGroupShapeProperties.NonVisualDrawingProperties == null)
                    {
                        group.NonVisualGroupShapeProperties.NonVisualDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    }
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
                                Console.WriteLine($"[DEBUG] CopyFooterFromMaster: Adding nested ImagePart to SlidePart with ContentType: '{originalImagePart.ContentType}' and explicit ID.");
                                string newRelId = "rId" + Guid.NewGuid().ToString("N");
                                ImagePart newImagePart = slidePart.AddNewPart<ImagePart>(originalImagePart.ContentType, newRelId); 
                                using (var stream = originalImagePart.GetStream())
                                {
                                    stream.CopyTo(newImagePart.GetStream());
                                }
                                descendantBlip.Embed = newRelId; // Use the explicitly generated ID
                                Console.WriteLine($"[DEBUG] Imagem aninhada em grupo com ID '{originalImageRelId}' copiada e RelationshipId atualizado.");
                            }
                        }
                    }
                }
                else if (clonedElement is DocumentFormat.OpenXml.Presentation.GraphicFrame graphicFrame)
                {
                    string uniqueName = "ClonedFooterGraphicFrame_" + Guid.NewGuid().ToString();
                    graphicFrame.NonVisualGraphicFrameProperties ??= new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameProperties();
                    if (graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties == null)
                    {
                        graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties();
                    }
                    graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = new StringValue(uniqueName);

                    var blip = clonedElement.Descendants<D.Blip>().FirstOrDefault();
                    if (blip?.Embed?.Value != null)
                    {
                        string originalImageRelId = blip.Embed.Value;
                        ImagePart originalImagePart = (ImagePart)masterPart.GetPartById(originalImageRelId);

                        if (originalImagePart != null)
                        {
                            Console.WriteLine($"[DEBUG] CopyFooterFromMaster: Adding GraphicFrame ImagePart to SlidePart with ContentType: '{originalImagePart.ContentType}' and explicit ID.");
                            string newRelId = "rId" + Guid.NewGuid().ToString("N");
                            ImagePart newImagePart = slidePart.AddNewPart<ImagePart>(originalImagePart.ContentType, newRelId); 
                            using (var stream = originalImagePart.GetStream())
                            {
                                stream.CopyTo(newImagePart.GetStream());
                            }
                            blip.Embed = newRelId; // Use the explicitly generated ID
                            Console.WriteLine($"[DEBUG] SVG/Imagem '{GetElementName(graphicFrame)}' copiado e RelationshipId atualizado para '{newRelId}'.");
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
            Console.WriteLine($"[DEBUG] Chamando CopyFooterFromMaster para o slide {slidePart.Uri.OriginalString}.");
            // The destinationDocument parameter was removed as it's not directly used in CopyFooterFromMaster anymore
            CopyFooterFromMaster(slidePart, masterPart, footerLimitY);
        }

        // Helper method to update relationship IDs within the XML of a copied part
        private static void UpdatePartRelationships(OpenXmlPart newPart, Dictionary<string, string> oldToNewRelIdMap)
        {
            OpenXmlElement? rootElement = null;
            if (newPart is SlidePart slidePart)
            {
                rootElement = slidePart.Slide;
            }
            else if (newPart is SlideMasterPart slideMasterPart)
            {
                rootElement = slideMasterPart.SlideMaster;
            }
            else if (newPart is SlideLayoutPart slideLayoutPart)
            {
                rootElement = slideLayoutPart.SlideLayout;
            }

            if (rootElement == null) return;

            // Collect changes to apply after iteration to avoid modifying collection while iterating
            var changes = new List<(OpenXmlElement element, OpenXmlAttribute oldAttr, OpenXmlAttribute newAttr)>();

            foreach (var element in rootElement.Descendants())
            {
                var attributesToChange = new List<OpenXmlAttribute>();
                foreach (var attribute in element.GetAttributes())
                {
                    // VERIFICAÇÃO ALTERADA: Usando a string literal do NamespaceUri
                    if (attribute.LocalName == "id" && attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                    {
                        if (oldToNewRelIdMap.TryGetValue(attribute.Value!, out string? newRelId))
                        {
                            attributesToChange.Add(attribute);
                            changes.Add((element, attribute, new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newRelId)));
                        }
                    }
                }
            }

            // Apply collected changes
            foreach (var change in changes)
            {
                change.element.SetAttribute(change.newAttr);
                Console.WriteLine($"[DEBUG] Updated r:id from '{change.oldAttr.Value}' to '{change.newAttr.Value}' for element '{change.element.LocalName}' in part '{newPart.Uri.OriginalString}'.");
            }
        }

        public async Task<SlideMergeResponse> MergeSlides(IFormFile destinationFile, SlideMergeRequest request)
        {
            Console.WriteLine("[DEBUG] Iniciando MergeSlides.");
            try
            {
                Console.WriteLine("[DEBUG] Criando caminhos temporários.");
                string tempDestinationPath = System.IO.Path.GetTempFileName() + ".pptx";
                string outputPath = System.IO.Path.GetTempFileName() + ".pptx";

                Console.WriteLine("[DEBUG] Copiando arquivo de destino para caminho temporário.");
                using (var stream = new System.IO.FileStream(tempDestinationPath, System.IO.FileMode.Create))
                {
                    await destinationFile.CopyToAsync(stream);
                }

                Console.WriteLine("[DEBUG] Copiando tempDestinationPath para outputPath.");
                System.IO.File.Copy(tempDestinationPath, outputPath, true);

                Console.WriteLine("[DEBUG] Abrindo documentos de apresentação.");
                using (var destino = PresentationDocument.Open(outputPath, true))
                using (var origem = PresentationDocument.Open(TemplatePath, false))
                {
                    Console.WriteLine("[DEBUG] Obtendo PresentationParts.");
                    var destinoPres = destino.PresentationPart!;
                    var origemPres = origem.PresentationPart!;
                    var destinoSlides = destinoPres.Presentation!.SlideIdList!;
                    var origemSlideIds = origemPres.Presentation!.SlideIdList!.Elements<SlideId>().ToList();

                    int[] primeiros = { 0, 1 };
                    int ultimoIdx = origemSlideIds.Count - 1;

                    uint NextSlideId() =>
                        destinoSlides.Elements<SlideId>().Any() ? 
                        destinoSlides.Elements<SlideId>().Max(s => s.Id!.Value) + 1 : 256U;

                    Console.WriteLine("[DEBUG] Processando primeiros slides.");
                    foreach (int i in primeiros)
                    {
                        var origemSlidePart = (SlidePart)origemPres.GetPartById(origemSlideIds[i].RelationshipId!);

                        string marcadorVerificacao = i == 0 ? "Prof" : "Lei nº 9610/98";

                        if (SlideWithTextExists(destinoPres, marcadorVerificacao))
                        {
                            Console.WriteLine($"[DEBUG] Slide com marcador '{marcadorVerificacao}' já existe. Pulando.");
                            continue;
                        }
                        
                        string newSlideRelId = "rId" + Guid.NewGuid().ToString("N");
                        Console.WriteLine($"[DEBUG] AddNewPart SlidePart com ContentType: '{origemSlidePart.ContentType}' e explicit ID: '{newSlideRelId}'");
                        SlidePart novoSlidePart = destinoPres.AddNewPart<SlidePart>(origemSlidePart.ContentType, newSlideRelId);
                        var slideRelIdMap = new Dictionary<string, string>(); // Map for this slide's internal relationships

                        using (Stream stream = origemSlidePart.GetStream(FileMode.Open))
                        {
                            novoSlidePart.FeedData(stream);
                        }

                        foreach (var partPair in origemSlidePart.Parts) 
                        {
                            OpenXmlPart relatedPart = partPair.OpenXmlPart;

                            if (relatedPart is ImagePart originalImagePart)
                            {
                                string newImageRelId = "rId" + Guid.NewGuid().ToString("N");
                                Console.WriteLine($"[DEBUG] AddNewPart ImagePart with ContentType: '{originalImagePart.ContentType}' e explicit ID: '{newImageRelId}'");
                                ImagePart newImagePart = novoSlidePart.AddNewPart<ImagePart>(originalImagePart.ContentType, newImageRelId); 
                                using (Stream imageStream = originalImagePart.GetStream(FileMode.Open))
                                {
                                    newImagePart.FeedData(imageStream);
                                }
                                slideRelIdMap[partPair.RelationshipId] = newImageRelId;
                            }
                            // Add other part types if necessary (e.g., ChartPart, DiagramColorsPart)
                        }
                        UpdatePartRelationships(novoSlidePart, slideRelIdMap); // Apply relationship updates

                        uint novoId = NextSlideId();
                        ReplaceTextColor(novoSlidePart, "NOMEMBA", request.Theme);
                        ReplaceTextInSlide(novoSlidePart, "NOMEMBA", request.Mba.ToUpper());
                        ReplaceTextInSlide(novoSlidePart, "Título da aula/disciplina", request.TituloAula);
                        ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");
                        ReplaceTextArea(novoSlidePart, "Lei nº 9610/98", request.Theme);

                        destinoSlides.InsertAt(new SlideId
                        {
                            Id = novoId,
                            RelationshipId = newSlideRelId // Use the explicitly generated ID here
                        }, i);
                        Console.WriteLine($"[DEBUG] Slide {i} adicionado com sucesso.");
                    }

                    Console.WriteLine("[DEBUG] Processando slide final (linkedin).");
                    if (!SlideWithTextExists(destinoPres, "linkedin.com/in/"))
                    {
                        var origemSlidePart = (SlidePart)origemPres.GetPartById(origemSlideIds[ultimoIdx].RelationshipId!);
                        
                        string newFinalSlideRelId = "rId" + Guid.NewGuid().ToString("N");
                        Console.WriteLine($"[DEBUG] AddNewPart SlidePart (final) com ContentType: '{origemSlidePart.ContentType}' e explicit ID: '{newFinalSlideRelId}'");
                        SlidePart novoSlidePart = destinoPres.AddNewPart<SlidePart>(origemSlidePart.ContentType, newFinalSlideRelId);
                        var slideRelIdMap = new Dictionary<string, string>(); // Map for this slide's internal relationships

                        using (Stream stream = origemSlidePart.GetStream(FileMode.Open))
                        {
                            novoSlidePart.FeedData(stream);
                        }

                        foreach (var partPair in origemSlidePart.Parts) 
                        {
                            OpenXmlPart relatedPart = partPair.OpenXmlPart;
                            if (relatedPart is ImagePart originalImagePart)
                            {
                                string newImageRelId = "rId" + Guid.NewGuid().ToString("N");
                                Console.WriteLine($"[DEBUG] AddNewPart ImagePart (final) com ContentType: '{originalImagePart.ContentType}' e explicit ID: '{newImageRelId}'");
                                ImagePart newImagePart = novoSlidePart.AddNewPart<ImagePart>(originalImagePart.ContentType, newImageRelId);
                                using (Stream imageStream = originalImagePart.GetStream(FileMode.Open))
                                {
                                    newImagePart.FeedData(imageStream);
                                }
                                slideRelIdMap[partPair.RelationshipId] = newImageRelId;
                            }
                            // Add other part types if necessary
                        }
                        UpdatePartRelationships(novoSlidePart, slideRelIdMap); // Apply relationship updates

                        ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");
                        ReplaceTextColor(novoSlidePart, "linkedin.perfil.com", request.Theme);
                        ReplaceTextInSlide(novoSlidePart, "linkedin.perfil.com", request.LinkedinPerfil);

                        destinoSlides.Append(new SlideId
                        {
                            Id = NextSlideId(),
                            RelationshipId = newFinalSlideRelId // Use the explicitly generated ID here
                        });
                        Console.WriteLine("[DEBUG] Slide final adicionado com sucesso.");
                    }

                    Console.WriteLine("[DEBUG] Processando Slide Master e Layouts.");
                    var thirdSlide = (SlidePart)origemPres.GetPartById(origemSlideIds[2].RelationshipId!);
                    var layoutSrc = thirdSlide.SlideLayoutPart!;
                    var masterSrc = layoutSrc.SlideMasterPart!;
                    var layoutName = layoutSrc.SlideLayout.Type?.Value.ToString() ?? "default";

                    var existingMaster = destinoPres.SlideMasterParts
                        .FirstOrDefault(mp => mp.SlideMaster?.Descendants<SlideLayout>()
                            .Any(l => (l.Type?.Value.ToString() ?? "default") == layoutName) ?? false);

                    SlideMasterPart masterDest;
                    SlideLayoutPart? layoutDest = null; 

                    if (existingMaster == null)
                    {
                        Console.WriteLine("[DEBUG] Slide Master não existente. Copiando do template.");
                        string newMasterRelId = "rId" + Guid.NewGuid().ToString("N");
                        Console.WriteLine($"[DEBUG] AddNewPart SlideMasterPart com ContentType: '{masterSrc.ContentType}' e explicit ID: '{newMasterRelId}'");
                        masterDest = destinoPres.AddNewPart<SlideMasterPart>(masterSrc.ContentType, newMasterRelId);
                        var masterRelIdMap = new Dictionary<string, string>();

                        using (var stream = masterSrc.GetStream(FileMode.Open))
                        {
                            masterDest.FeedData(stream);
                        }
                        
                        foreach (var partPair in masterSrc.Parts) 
                        {
                            OpenXmlPart relatedMasterSrcPart = partPair.OpenXmlPart;
                            if (relatedMasterSrcPart is ThemePart originalThemePart)
                            {
                                string newThemeRelId = "rId" + Guid.NewGuid().ToString("N");
                                Console.WriteLine($"[DEBUG] AddNewPart ThemePart com ContentType: '{originalThemePart.ContentType}' e explicit ID: '{newThemeRelId}'");
                                ThemePart newThemePart = masterDest.AddNewPart<ThemePart>(originalThemePart.ContentType, newThemeRelId);
                                using (Stream themeStream = originalThemePart.GetStream(FileMode.Open))
                                {
                                    newThemePart.FeedData(themeStream);
                                }
                                masterRelIdMap[partPair.RelationshipId] = newThemeRelId;
                            }
                            else if (relatedMasterSrcPart is FontPart originalFontPart) 
                            {
                                string newFontRelId = "rId" + Guid.NewGuid().ToString("N");
                                Console.WriteLine($"[DEBUG] AddNewPart FontPart com ContentType: '{originalFontPart.ContentType}' e explicit ID: '{newFontRelId}'");
                                FontPart newFontPart = masterDest.AddNewPart<FontPart>(originalFontPart.ContentType, newFontRelId);
                                using (Stream fontStream = originalFontPart.GetStream(FileMode.Open))
                                {
                                    newFontPart.FeedData(fontStream);
                                }
                                masterRelIdMap[partPair.RelationshipId] = newFontRelId;
                            }
                            // Add other part types if necessary
                        }
                        UpdatePartRelationships(masterDest, masterRelIdMap); // Apply relationship updates to master

                        foreach (var laySrcPart in masterSrc.SlideLayoutParts) 
                        {
                            string newLayoutRelId = "rId" + Guid.NewGuid().ToString("N");
                            Console.WriteLine($"[DEBUG] AddNewPart SlideLayoutPart com ContentType: '{laySrcPart.ContentType}' e explicit ID: '{newLayoutRelId}'");
                            SlideLayoutPart newLayoutPart = masterDest.AddNewPart<SlideLayoutPart>(laySrcPart.ContentType, newLayoutRelId);
                            var layoutRelIdMap = new Dictionary<string, string>();
                            using (Stream stream = laySrcPart.GetStream(FileMode.Open))
                            {
                                newLayoutPart.FeedData(stream);
                            }

                            foreach (var partPairLayout in laySrcPart.Parts) 
                            {
                                OpenXmlPart relatedLayoutSrcPart = partPairLayout.OpenXmlPart;
                                if (relatedLayoutSrcPart is ImagePart originalLayoutImagePart)
                                {
                                    string newLayoutImageRelId = "rId" + Guid.NewGuid().ToString("N");
                                    Console.WriteLine($"[DEBUG] AddNewPart ImagePart (layout) com ContentType: '{originalLayoutImagePart.ContentType}' e explicit ID: '{newLayoutImageRelId}'");
                                    ImagePart newLayoutImagePart = newLayoutPart.AddNewPart<ImagePart>(originalLayoutImagePart.ContentType, newLayoutImageRelId);
                                    using (Stream imgStream = originalLayoutImagePart.GetStream(FileMode.Open))
                                    {
                                        newLayoutImagePart.FeedData(imgStream);
                                    }
                                    layoutRelIdMap[partPairLayout.RelationshipId] = newLayoutImageRelId;
                                }
                                // Add other part types if necessary
                            }
                            UpdatePartRelationships(newLayoutPart, layoutRelIdMap); // Apply relationship updates to layout
                            
                            if ((newLayoutPart.SlideLayout.Type?.Value.ToString() ?? "default") == layoutName)
                            {
                                layoutDest = newLayoutPart;
                                Console.WriteLine($"[DEBUG] Layout '{layoutName}' encontrado e definido como destino.");
                            }
                        }

                        var masterIdList = destinoPres.Presentation!.SlideMasterIdList 
                                                     ?? destinoPres.Presentation.AppendChild(new SlideMasterIdList());

                        uint nextMasterId = masterIdList.Elements<SlideMasterId>().Any()
                            ? masterIdList.Elements<SlideMasterId>().Max(m => m.Id!.Value) + 1
                            : 1;

                        masterIdList.Append(new SlideMasterId
                        {
                            Id = nextMasterId,
                            RelationshipId = newMasterRelId // Use the explicitly generated ID here
                        });
                        Console.WriteLine("[DEBUG] Slide Master adicionado ao PresentationPart de destino.");
                    }
                    else
                    {
                        masterDest = existingMaster;
                        layoutDest = masterDest.SlideLayoutParts
                            .First(l => (l.SlideLayout.Type?.Value.ToString() ?? "default") == layoutName);
                        Console.WriteLine("[DEBUG] Usando Slide Master existente.");
                    }

                    var currentSlideIdsInDest = destinoSlides.Elements<SlideId>().ToList();
                    
                    var footerLimitY = GetFooterStartY(destino);

                    Console.WriteLine("[DEBUG] Iterando sobre slides para aplicar layout e rodapé.");
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
                        
                        if (layoutDest != null)
                        {
                            slidePart.AddPart(layoutDest);
                        }
                        else
                        {
                            Console.WriteLine($"[WARNING] layoutDest é nulo para o slide {slidePart.Uri.OriginalString}. Layout não aplicado.");
                        }
                        slidePart.Slide!.Save(); 

                        // Always call AddFooterOverlayToSlide for these slides
                        Console.WriteLine($"[DEBUG] Aplicando rodapé ao slide {slidePart.Uri.OriginalString} (index {i}).");
                        AddFooterOverlayToSlide(slidePart, masterSrc, footerLimitY);
                        
                        // If there is content overlapping the footer, you might want to adjust it here if needed.
                        // The HasContentOverlappingFooter can still be used for logging or future adjustment logic.
                        if (HasContentOverlappingFooter(slidePart, footerLimitY))
                        {
                            Console.WriteLine($"[DEBUG] Conteúdo sobreposto ao rodapé detectado no slide {slidePart.Uri.OriginalString}. Considere ajustar o conteúdo existente.");
                        }
                    }
                    
                    Console.WriteLine("[DEBUG] Aplicando numeração de página.");
                    ApplyPageNumbering(destinoPres);
                    Console.WriteLine("[DEBUG] Salvando apresentação de destino.");
                    destinoPres.Presentation!.Save(); 
                }

                Console.WriteLine("[DEBUG] Deletando arquivos temporários.");
                System.IO.File.Delete(tempDestinationPath);

                string outputNormalizedPath = System.IO.Path.GetTempFileName() + ".pptx";
                Console.WriteLine("[DEBUG] Normalizando com Python.");
                await NormalizarComPythonAsync(outputPath, outputNormalizedPath);

                System.IO.File.Delete(outputPath);

                var fileName = System.IO.Path.GetFileName(outputNormalizedPath);
                var downloadUrl = outputNormalizedPath;

                Console.WriteLine("[DEBUG] Processamento concluído com sucesso.");
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
                Console.WriteLine($"[ERROR] Exceção capturada: {ex.Message}");
                Console.WriteLine($"[ERROR] Stack Trace: {ex.StackTrace}");
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