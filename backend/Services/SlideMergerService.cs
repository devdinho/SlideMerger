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
            Console.WriteLine("[DEBUG] Tentando aplicar numeração de página via configurações da apresentação.");

            // Obtém ou cria o PresentationPropertiesPart
            PresentationPropertiesPart? presentationPropertiesPart = presentationPart.PresentationPropertiesPart;
            if (presentationPropertiesPart == null)
            {
                presentationPropertiesPart = presentationPart.AddNewPart<PresentationPropertiesPart>();
                presentationPropertiesPart.PresentationProperties = new DocumentFormat.OpenXml.Presentation.PresentationProperties(); // Explicitly use the PresentationProperties type
                Console.WriteLine("[DEBUG] PresentationPropertiesPart criado.");
            }

            // Obtém ou cria o elemento HeaderFooter dentro de PresentationProperties
            // HeaderFooter é uma classe dentro de DocumentFormat.OpenXml.Presentation
            DocumentFormat.OpenXml.Presentation.HeaderFooter? headerFooter = presentationPropertiesPart.PresentationProperties.GetFirstChild<DocumentFormat.OpenXml.Presentation.HeaderFooter>();
            if (headerFooter == null)
            {
                headerFooter = new DocumentFormat.OpenXml.Presentation.HeaderFooter();
                presentationPropertiesPart.PresentationProperties.AppendChild(headerFooter);
                Console.WriteLine("[DEBUG] HeaderFooter element criado dentro de PresentationProperties.");
            }

            // Define a visibilidade do número do slide. BooleanValue está em DocumentFormat.OpenXml
            // A propriedade correta é 'SlideNumber', não 'SlideNumberVisibility'
            if (headerFooter.SlideNumber == null)
            {
                headerFooter.SlideNumber = new DocumentFormat.OpenXml.BooleanValue(true);
            }
            else
            {
                headerFooter.SlideNumber.Value = true;
            }
            Console.WriteLine("[DEBUG] Propriedade de visibilidade do número do slide definida como TRUE no HeaderFooter do PresentationPropertiesPart.");

            // Salva as mudanças no PresentationPropertiesPart
            presentationPropertiesPart.PresentationProperties.Save();
            Console.WriteLine("[DEBUG] PresentationPropertiesPart salvo.");

            // Isso apenas habilita a *exibição* dos números de slide.
            // Para que os números realmente apareçam, seu Slide Mestre(s) e/ou Layout(s) de Slide
            // no arquivo Template.pptx devem ter o placeholder nativo "Número de Slide" (Type: sldNum).
            // Se não tiverem, ou se estiverem ocultos no mestre/layout, eles ainda não aparecerão.
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


        public static void CopyFooterFromMasterScaled(
            SlidePart slidePart,
            SlideMasterPart masterPart,
            long footerYStart,
            long templateCx,
            long templateCy,
            long targetCx,
            long targetCy)
        {
            double scaleX = targetCx / (double)templateCx;
            double scaleY = targetCy / (double)templateCy;

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var footerElementsToCopy = new List<OpenXmlElement>();

            var allElements = masterPart.SlideMaster.CommonSlideData.ShapeTree.Elements<OpenXmlElement>();

            foreach (var element in allElements)
            {
                bool isGroup = element is DocumentFormat.OpenXml.Presentation.GroupShape;
                bool shouldCopy = false;

                long? yPos = GetYPosition(element);

                if (isGroup)
                {
                    var children = element.Descendants<OpenXmlElement>()
                        .Where(e => e is DocumentFormat.OpenXml.Presentation.Shape || e is D.Picture).ToList();

                    foreach (var child in children)
                    {
                        var childY = GetYPosition(child);
                        if (childY.HasValue && childY >= footerYStart)
                        {
                            shouldCopy = true;
                            break;
                        }
                    }
                }
                else if (yPos.HasValue && yPos >= footerYStart)
                {
                    shouldCopy = true;
                }

                if (shouldCopy)
                {
                    footerElementsToCopy.Add(element);
                }
            }

            foreach (var footerElement in footerElementsToCopy)
            {
                var clonedElement = (OpenXmlElement)footerElement.CloneNode(true);

                // Aplica escala a elementos individuais
                void ScaleTransform(OpenXmlElement elem)
                {
                    var t2d = elem.Descendants<Transform2D>().FirstOrDefault();
                    if (t2d?.Offset != null)
                    {
                        t2d.Offset.X = (long)(t2d.Offset.X * scaleX);
                        t2d.Offset.Y = (long)(t2d.Offset.Y * scaleY);
                    }
                    if (t2d?.Extents != null)
                    {
                        t2d.Extents.Cx = (long)(t2d.Extents.Cx * scaleX);
                        t2d.Extents.Cy = (long)(t2d.Extents.Cy * scaleY);
                    }
                }

                // Aplica escala recursiva em GroupShape
                void ScaleGroup(OpenXmlElement group)
                {
                    foreach (var child in group.Elements())
                    {
                        ScaleTransform(child);
                        if (child is DocumentFormat.OpenXml.Presentation.GroupShape subgroup)
                        {
                            ScaleGroup(subgroup);
                        }
                    }
                }

                ScaleTransform(clonedElement);
                if (clonedElement is DocumentFormat.OpenXml.Presentation.GroupShape)
                    ScaleGroup(clonedElement);

                // Corrigir imagens (Blips)
                var blips = clonedElement.Descendants<D.Blip>();
                foreach (var blip in blips)
                {
                    string? relId = blip.Embed?.Value;
                    if (relId == null) continue;

                    ImagePart? sourceImagePart = masterPart.GetPartById(relId) as ImagePart;
                    if (sourceImagePart == null) continue;

                    string? newRelId;
                    if (!slidePart.Parts.Any(p => p.OpenXmlPart == sourceImagePart))
                    {
                        var newImagePart = slidePart.AddImagePart(sourceImagePart.ContentType);
                        using var imgStream = sourceImagePart.GetStream();
                        newImagePart.FeedData(imgStream);
                        newRelId = slidePart.GetIdOfPart(newImagePart);
                    }
                    else
                    {
                        newRelId = slidePart.GetIdOfPart(sourceImagePart);
                    }

                    if (newRelId != null)
                        blip.Embed = newRelId;
                }

                // Renomear para evitar conflitos
                if (clonedElement is DocumentFormat.OpenXml.Presentation.Shape shape)
                {
                    shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterShape_" + Guid.NewGuid().ToString());
                }
                else if (clonedElement is D.Picture pic)
                {
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterPicture_" + Guid.NewGuid().ToString());
                }
                else if (clonedElement is DocumentFormat.OpenXml.Presentation.GroupShape group)
                {
                    group.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Name =
                        new StringValue("ClonedFooterGroup_" + Guid.NewGuid().ToString());
                }

                shapeTree.Append(clonedElement);
                Console.WriteLine($"[DEBUG] Rodapé '{GetElementName(clonedElement)}' copiado e escalado para {slidePart.Uri.OriginalString}");
            }

            slidePart.Slide.Save();
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

        public static SlidePart? GetSlidePartWithText(PresentationPart presPart, string searchText)
        {
            foreach (var sp in presPart.SlideParts)
            {
                if (sp.Slide?.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                        .Any(t => t.Text?.Contains(searchText) == true) ?? false)
                {
                    return sp;
                }
            }
            return null;
        }

        public static string? GetTextAreaColor(SlidePart slidePart, string textToFind)
        {
            if (slidePart.Slide == null) return null;

            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();

            foreach (var shape in shapes)
            {
                var text = shape.TextBody?.InnerText;
                if (!string.IsNullOrEmpty(text) && text.Contains(textToFind))
                {
                    var solidFill = shape.ShapeProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
                    if (solidFill != null)
                    {
                        var rgbColor = solidFill.GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>();
                        if (rgbColor != null)
                        {
                            return rgbColor.Val?.Value; // Retorna a cor em formato hexadecimal (ex: "FF0000")
                        }
                    }
                }
            }
            return null;
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
                        string marcadorVerificacao = i == 0 ? $"{request.Mba.ToUpper()}" : "Lei nº 9610/98";

                        bool shouldAddFromTemplate = true;
                        SlidePart? existingSlidePart = GetSlidePartWithText(destinoPres, marcadorVerificacao); // Usando o novo método

                        if (existingSlidePart != null)
                        {
                            if (i == 1) // Específico para o slide da Lei nº 9610/98
                            {
                                string? existingColor = GetTextAreaColor(existingSlidePart, "Lei nº 9610/98"); // Usando o novo método
                                string targetColor = request.Theme.Replace("#", "");

                                if (existingColor != null && existingColor.Equals(targetColor, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"[DEBUG] Slide 'Lei nº 9610/98' existe e a cor da área ('{existingColor}') corresponde à cor do tema ('{targetColor}'). Pulando adição do template.");
                                    shouldAddFromTemplate = false;
                                }
                                else
                                {
                                    Console.WriteLine($"[DEBUG] Slide 'Lei nº 9610/98' existe, mas a cor da área ('{existingColor ?? "N/A"}') NÃO corresponde à cor do tema ('{targetColor}'). Removendo slide existente.");
                                    // Remove o slide existente
                                    var slideIdToRemove = destinoSlides.Elements<SlideId>().FirstOrDefault(sid => destinoPres.GetPartById(sid.RelationshipId!) == existingSlidePart);
                                    if (slideIdToRemove != null)
                                    {
                                        destinoSlides.RemoveChild(slideIdToRemove);
                                        destinoPres.DeletePart(existingSlidePart);
                                        Console.WriteLine($"[DEBUG] Slide 'Lei nº 9610/98' existente removido.");
                                    }
                                }
                            }
                            else // Para o slide "Prof" (i == 0) ou qualquer outro que não seja a Lei 9610/98
                            {
                                Console.WriteLine($"[DEBUG] Slide com marcador '{marcadorVerificacao}' já existe. Pulando.");
                                shouldAddFromTemplate = false;
                            }
                        }
                        else
                        {
                            Console.WriteLine($"[DEBUG] Slide com marcador '{marcadorVerificacao}' não existe no destino. Adicionando do template.");
                        }

                        if (shouldAddFromTemplate)
                        {
                            string newSlideRelId = "rId" + Guid.NewGuid().ToString("N");
                            Console.WriteLine($"[DEBUG] AddNewPart SlidePart com ContentType: '{origemSlidePart.ContentType}' e explicit ID: '{newSlideRelId}'");
                            SlidePart novoSlidePart = destinoPres.AddNewPart<SlidePart>(origemSlidePart.ContentType, newSlideRelId);
                            var slideRelIdMap = new Dictionary<string, string>();

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
                            }
                            UpdatePartRelationships(novoSlidePart, slideRelIdMap);

                            uint novoId = NextSlideId();

                            // Aplica as substituições de texto e cor ao slide recém-adicionado
                            if (i == 0) // Slide "Prof"
                            {
                                ReplaceTextColor(novoSlidePart, "NOMEMBA", request.Theme);
                                ReplaceTextInSlide(novoSlidePart, "NOMEMBA", request.Mba.ToUpper());
                                ReplaceTextInSlide(novoSlidePart, "Título da aula/disciplina", request.TituloAula);
                                ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");
                            }
                            else if (i == 1) // Slide "Lei nº 9610/98"
                            {
                                ReplaceTextArea(novoSlidePart, "Lei nº 9610/98", request.Theme); // Garante a cor correta para o novo slide
                                // Outras substituições de texto se houverem
                            }


                            destinoSlides.InsertAt(new SlideId
                            {
                                Id = novoId,
                                RelationshipId = newSlideRelId
                            }, i);
                            Console.WriteLine($"[DEBUG] Slide {i} adicionado com sucesso do template.");
                        }
                    }

                    Console.WriteLine("[DEBUG] Processando slide final (linkedin).");

                    bool temSlideProfessor = SlideWithTextExists(destinoPres, $"Prof(a) {request.NomeProfessor}");
                    bool temSlideLinkedin = SlideWithTextExists(destinoPres, request.LinkedinPerfil);
                    bool temSlideObrigado = SlideWithTextExists(destinoPres, "Obrigado(a)!");

                    if (temSlideProfessor && temSlideLinkedin && temSlideObrigado)
                    {
                        Console.WriteLine("[DEBUG] Slide final já existe. Pulando adição.");
                    }
                    else
                    {
                        if (temSlideProfessor || temSlideLinkedin || temSlideObrigado)
                        {
                            Console.WriteLine("[DEBUG] Slide final incompleto detectado. Removendo instância parcial.");
                            var slideIdToRemove = destinoSlides.Elements<SlideId>()
                                .LastOrDefault(); // ou use uma lógica mais precisa se não for sempre o último
                            if (slideIdToRemove != null)
                            {
                                destinoSlides.RemoveChild(slideIdToRemove);
                            }
                        }

                        // Agora sim adiciona o slide completo
                        var origemSlidePart = (SlidePart)origemPres.GetPartById(origemSlideIds[ultimoIdx].RelationshipId!);

                        string newFinalSlideRelId = "rId" + Guid.NewGuid().ToString("N");
                        Console.WriteLine($"[DEBUG] AddNewPart SlidePart (final) com ContentType: '{origemSlidePart.ContentType}' e explicit ID: '{newFinalSlideRelId}'");
                        SlidePart novoSlidePart = destinoPres.AddNewPart<SlidePart>(origemSlidePart.ContentType, newFinalSlideRelId);
                        var slideRelIdMap = new Dictionary<string, string>();

                        using (Stream stream = origemSlidePart.GetStream(FileMode.Open))
                        {
                            novoSlidePart.FeedData(stream);
                        }

                        foreach (var partPair in origemSlidePart.Parts)
                        {
                            if (partPair.OpenXmlPart is ImagePart originalImagePart)
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
                        }

                        UpdatePartRelationships(novoSlidePart, slideRelIdMap);

                        ReplaceTextInSlide(novoSlidePart, "Nome do(a) Professor(a)", $"Prof(a) {request.NomeProfessor}");
                        ReplaceTextColor(novoSlidePart, "linkedin.perfil.com", request.Theme);
                        ReplaceTextInSlide(novoSlidePart, "linkedin.perfil.com", request.LinkedinPerfil);

                        destinoSlides.Append(new SlideId
                        {
                            Id = NextSlideId(),
                            RelationshipId = newFinalSlideRelId
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
                    long targetCx = destinoPres.Presentation.SlideSize.Cx;
                    long targetCy = destinoPres.Presentation.SlideSize.Cy;
                    long templateCx = origemPres.Presentation.SlideSize.Cx;
                    long templateCy = origemPres.Presentation.SlideSize.Cy;
                    long footerY = GetFooterStartY(destino);

                    Console.WriteLine("[DEBUG] Iterando sobre slides para aplicar layout e rodapé.");
                    for (int i = 0; i < currentSlideIdsInDest.Count; i++)
                    {
                        var slidePart = (SlidePart)destinoPres.GetPartById(currentSlideIdsInDest[i].RelationshipId!);

                        if (i < 2 || i == currentSlideIdsInDest.Count - 1)
                        {
                            Console.WriteLine($"[DEBUG] Ignorando aplicação de layout ou rodapé para slide index {i}.");
                            continue;
                        }


                        slidePart.Slide!.Save();

                        Console.WriteLine($"[DEBUG] Aplicando rodapé ao slide {slidePart.Uri.OriginalString} (index {i}).");
                        CopyFooterFromMasterScaled(slidePart, masterSrc, footerY, templateCx, templateCy, targetCx, targetCy);


                        if (HasContentOverlappingFooter(slidePart, footerLimitY))
                        {
                            Console.WriteLine($"[DEBUG] Conteúdo sobreposto ao rodapé detectado no slide {slidePart.Uri.OriginalString}. Considere ajustar o conteúdo existente.");
                        }
                    }

                    Console.WriteLine("[DEBUG] Aplicando numeração de página.");
                    // ApplyPageNumbering(destinoPres); // Esta chamada agora apenas garante a visibilidade geral do número do slide.
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