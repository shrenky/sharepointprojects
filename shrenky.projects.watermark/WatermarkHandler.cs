using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Vml;

namespace shrenky.projects.watermark
{
    public class WatermarkHandler
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private WordprocessingDocument document;
        private string WatermarkText;

        public Stream DocumentStream 
        {
            get 
            {
                if (document == null)
                {
                    return null;
                }
                else
                {
                    return document.MainDocumentPart.GetStream();
                }
            }
        }

        public void ChangePackage()
        {
            using (document)
            {
                ChangeParts();
            }
        }

        public WatermarkHandler(Stream stream, string text)
        {
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);
            this.document = doc;
            this.WatermarkText = text;
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Changes the relationship ID of the parts.
            ReconfigureRelationshipID();
            //Adds new parts or new relationships.
            AddParts();
            //Changes the contents of the specified parts.
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeMainDocumentPart1(document.MainDocumentPart);
            ChangeDocumentSettingsPart1(((DocumentSettingsPart)UriPartDictionary["/word/settings.xml"]));
            ChangeStyleDefinitionsPart1(((StyleDefinitionsPart)UriPartDictionary["/word/styles.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        /// <summary>
        /// Changes the relationship ID of the parts in the source package to make sure these IDs are the same as those in the target package.
        /// To avoid the conflict of the relationship ID, a temporary ID is assigned first.        
        /// </summary>
        private void ReconfigureRelationshipID()
        {
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/theme/theme1.xml"], "generatedTmpID1");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/fontTable.xml"], "generatedTmpID2");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/theme/theme1.xml"], "rId13");
            document.MainDocumentPart.ChangeIdOfPart(UriPartDictionary["/word/fontTable.xml"], "rId12");
        }

        /// <summary>
        /// Adds new parts or new relationship between parts.
        /// </summary>
        private void AddParts()
        {
            try
            {
                //Generate new parts.
                FooterPart footerPart1 = document.MainDocumentPart.AddNewPart<FooterPart>("rId8");
                GenerateFooterPart1Content(footerPart1);

                HeaderPart headerPart1 = document.MainDocumentPart.AddNewPart<HeaderPart>("rId7");
                GenerateHeaderPart1Content(headerPart1);

                HeaderPart headerPart2 = document.MainDocumentPart.AddNewPart<HeaderPart>("rId6");
                GenerateHeaderPart2Content(headerPart2);

                FooterPart footerPart2 = document.MainDocumentPart.AddNewPart<FooterPart>("rId11");
                GenerateFooterPart2Content(footerPart2);

                EndnotesPart endnotesPart1 = document.MainDocumentPart.AddNewPart<EndnotesPart>("rId5");
                GenerateEndnotesPart1Content(endnotesPart1);

                HeaderPart headerPart3 = document.MainDocumentPart.AddNewPart<HeaderPart>("rId10");
                GenerateHeaderPart3Content(headerPart3);

                FootnotesPart footnotesPart1 = document.MainDocumentPart.AddNewPart<FootnotesPart>("rId4");
                GenerateFootnotesPart1Content(footnotesPart1);

                FooterPart footerPart3 = document.MainDocumentPart.AddNewPart<FooterPart>("rId9");
                GenerateFooterPart3Content(footerPart3);
            }
            catch { }

        }

        private void GenerateFooterPart1Content(FooterPart footerPart1)
        {
            Footer footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            footer1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);

            paragraph1.Append(paragraphProperties1);

            footer1.Append(paragraph1);

            footerPart1.Footer = footer1;
        }

        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties2.Append(paragraphStyleId2);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
            V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
            V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
            V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
            V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
            V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
            V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
            V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
            V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
            V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
            V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
            V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            formulas1.Append(formula13);
            formulas1.Append(formula14);
            V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
            V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

            V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

            shapeHandles1.Append(shapeHandle1);
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(textPath1);
            shapetype1.Append(shapeHandles1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject12513528", Style = "position:absolute;margin-left:0;margin-top:0;width:348.05pt;height:261pt;rotation:315;z-index:-251653120;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2051", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
            V.Fill fill1 = new V.Fill() { Opacity = ".5" };
            V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = WatermarkText };

            shape1.Append(fill1);
            shape1.Append(textPath2);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run1.Append(runProperties1);
            run1.Append(picture1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run1);

            header1.Append(paragraph2);

            headerPart1.Header = header1;
        }

        private void GenerateHeaderPart2Content(HeaderPart headerPart2)
        {
            Header header2 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            header2.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header2.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header2.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header2.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header2.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header2.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header2.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header2.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header2.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header2.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header2.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header2.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header2.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties3.Append(paragraphStyleId3);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);

            Picture picture2 = new Picture();

            V.Shapetype shapetype2 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

            V.Formulas formulas2 = new V.Formulas();
            V.Formula formula15 = new V.Formula() { Equation = "sum #0 0 10800" };
            V.Formula formula16 = new V.Formula() { Equation = "prod #0 2 1" };
            V.Formula formula17 = new V.Formula() { Equation = "sum 21600 0 @1" };
            V.Formula formula18 = new V.Formula() { Equation = "sum 0 0 @2" };
            V.Formula formula19 = new V.Formula() { Equation = "sum 21600 0 @3" };
            V.Formula formula20 = new V.Formula() { Equation = "if @0 @3 0" };
            V.Formula formula21 = new V.Formula() { Equation = "if @0 21600 @1" };
            V.Formula formula22 = new V.Formula() { Equation = "if @0 0 @2" };
            V.Formula formula23 = new V.Formula() { Equation = "if @0 @4 21600" };
            V.Formula formula24 = new V.Formula() { Equation = "mid @5 @6" };
            V.Formula formula25 = new V.Formula() { Equation = "mid @8 @5" };
            V.Formula formula26 = new V.Formula() { Equation = "mid @7 @8" };
            V.Formula formula27 = new V.Formula() { Equation = "mid @6 @7" };
            V.Formula formula28 = new V.Formula() { Equation = "sum @6 0 @5" };

            formulas2.Append(formula15);
            formulas2.Append(formula16);
            formulas2.Append(formula17);
            formulas2.Append(formula18);
            formulas2.Append(formula19);
            formulas2.Append(formula20);
            formulas2.Append(formula21);
            formulas2.Append(formula22);
            formulas2.Append(formula23);
            formulas2.Append(formula24);
            formulas2.Append(formula25);
            formulas2.Append(formula26);
            formulas2.Append(formula27);
            formulas2.Append(formula28);
            V.Path path2 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
            V.TextPath textPath3 = new V.TextPath() { On = true, FitShape = true };

            V.ShapeHandles shapeHandles2 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle2 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

            shapeHandles2.Append(shapeHandle2);
            Ovml.Lock lock2 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

            shapetype2.Append(formulas2);
            shapetype2.Append(path2);
            shapetype2.Append(textPath3);
            shapetype2.Append(shapeHandles2);
            shapetype2.Append(lock2);

            V.Shape shape2 = new V.Shape() { Id = "PowerPlusWaterMarkObject12513527", Style = "position:absolute;margin-left:0;margin-top:0;width:348.05pt;height:261pt;rotation:315;z-index:-251655168;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2050", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
            V.Fill fill2 = new V.Fill() { Opacity = ".5" };
            V.TextPath textPath4 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = WatermarkText };

            shape2.Append(fill2);
            shape2.Append(textPath4);

            picture2.Append(shapetype2);
            picture2.Append(shape2);

            run2.Append(runProperties2);
            run2.Append(picture2);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run2);

            header2.Append(paragraph3);

            headerPart2.Header = header2;
        }

        private void GenerateFooterPart2Content(FooterPart footerPart2)
        {
            Footer footer2 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            footer2.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer2.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer2.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer2.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer2.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer2.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer2.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer2.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer2.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer2.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer2.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer2.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer2.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties4.Append(paragraphStyleId4);

            paragraph4.Append(paragraphProperties4);

            footer2.Append(paragraph4);

            footerPart2.Footer = footer2;
        }

        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            endnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "003F318F", RsidParagraphProperties = "00422D22", RsidRunAdditionDefault = "003F318F" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties5.Append(spacingBetweenLines1);

            Run run3 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run3.Append(separatorMark1);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run3);

            endnote1.Append(paragraph5);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "003F318F", RsidParagraphProperties = "00422D22", RsidRunAdditionDefault = "003F318F" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties6.Append(spacingBetweenLines2);

            Run run4 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run4.Append(continuationSeparatorMark1);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run4);

            endnote2.Append(paragraph6);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        private void GenerateHeaderPart3Content(HeaderPart headerPart3)
        {
            Header header3 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            header3.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header3.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header3.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header3.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header3.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header3.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header3.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header3.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header3.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header3.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header3.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header3.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header3.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties7.Append(paragraphStyleId5);

            Run run5 = new Run();

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);

            Picture picture3 = new Picture();

            V.Shapetype shapetype3 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

            V.Formulas formulas3 = new V.Formulas();
            V.Formula formula29 = new V.Formula() { Equation = "sum #0 0 10800" };
            V.Formula formula30 = new V.Formula() { Equation = "prod #0 2 1" };
            V.Formula formula31 = new V.Formula() { Equation = "sum 21600 0 @1" };
            V.Formula formula32 = new V.Formula() { Equation = "sum 0 0 @2" };
            V.Formula formula33 = new V.Formula() { Equation = "sum 21600 0 @3" };
            V.Formula formula34 = new V.Formula() { Equation = "if @0 @3 0" };
            V.Formula formula35 = new V.Formula() { Equation = "if @0 21600 @1" };
            V.Formula formula36 = new V.Formula() { Equation = "if @0 0 @2" };
            V.Formula formula37 = new V.Formula() { Equation = "if @0 @4 21600" };
            V.Formula formula38 = new V.Formula() { Equation = "mid @5 @6" };
            V.Formula formula39 = new V.Formula() { Equation = "mid @8 @5" };
            V.Formula formula40 = new V.Formula() { Equation = "mid @7 @8" };
            V.Formula formula41 = new V.Formula() { Equation = "mid @6 @7" };
            V.Formula formula42 = new V.Formula() { Equation = "sum @6 0 @5" };

            formulas3.Append(formula29);
            formulas3.Append(formula30);
            formulas3.Append(formula31);
            formulas3.Append(formula32);
            formulas3.Append(formula33);
            formulas3.Append(formula34);
            formulas3.Append(formula35);
            formulas3.Append(formula36);
            formulas3.Append(formula37);
            formulas3.Append(formula38);
            formulas3.Append(formula39);
            formulas3.Append(formula40);
            formulas3.Append(formula41);
            formulas3.Append(formula42);
            V.Path path3 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
            V.TextPath textPath5 = new V.TextPath() { On = true, FitShape = true };

            V.ShapeHandles shapeHandles3 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle3 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

            shapeHandles3.Append(shapeHandle3);
            Ovml.Lock lock3 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

            shapetype3.Append(formulas3);
            shapetype3.Append(path3);
            shapetype3.Append(textPath5);
            shapetype3.Append(shapeHandles3);
            shapetype3.Append(lock3);

            V.Shape shape3 = new V.Shape() { Id = "PowerPlusWaterMarkObject12513526", Style = "position:absolute;margin-left:0;margin-top:0;width:348.05pt;height:261pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2049", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
            V.Fill fill3 = new V.Fill() { Opacity = ".5" };
            V.TextPath textPath6 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = WatermarkText };

            shape3.Append(fill3);
            shape3.Append(textPath6);

            picture3.Append(shapetype3);
            picture3.Append(shape3);

            run5.Append(runProperties3);
            run5.Append(picture3);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run5);

            header3.Append(paragraph7);

            headerPart3.Header = header3;
        }

        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            footnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "003F318F", RsidParagraphProperties = "00422D22", RsidRunAdditionDefault = "003F318F" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties8.Append(spacingBetweenLines3);

            Run run6 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run6.Append(separatorMark2);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run6);

            footnote1.Append(paragraph8);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "003F318F", RsidParagraphProperties = "00422D22", RsidRunAdditionDefault = "003F318F" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties9.Append(spacingBetweenLines4);

            Run run7 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run7.Append(continuationSeparatorMark2);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run7);

            footnote2.Append(paragraph9);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        private void GenerateFooterPart3Content(FooterPart footerPart3)
        {
            Footer footer3 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            footer3.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer3.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer3.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer3.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer3.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer3.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer3.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer3.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer3.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer3.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer3.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer3.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer3.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00422D22", RsidRunAdditionDefault = "00422D22" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties10.Append(paragraphStyleId6);

            paragraph10.Append(paragraphProperties10);

            footer3.Append(paragraph10);

            footerPart3.Footer = footer3;
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.TotalTime totalTime1 = properties1.GetFirstChild<Ap.TotalTime>();
            Ap.Company company1 = properties1.GetFirstChild<Ap.Company>();
            totalTime1.Text = "0";


            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Title";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);
            properties1.InsertBefore(headingPairs1, company1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            properties1.InsertBefore(titlesOfParts1, company1);
        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            OpenXmlPackage package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-11-24T01:46:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-11-24T01:46:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeMainDocumentPart1(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = mainDocumentPart1.Document;

            Body body1 = document1.GetFirstChild<Body>();

            Paragraph paragraph1 = body1.GetFirstChild<Paragraph>();
            Paragraph paragraph2 = body1.Elements<Paragraph>().ElementAt(3);
            Paragraph paragraph3 = body1.Elements<Paragraph>().ElementAt(4);
            Paragraph paragraph4 = body1.Elements<Paragraph>().ElementAt(5);
            Paragraph paragraph5 = body1.Elements<Paragraph>().ElementAt(6);
            Paragraph paragraph6 = body1.Elements<Paragraph>().ElementAt(7);
            Paragraph paragraph7 = body1.Elements<Paragraph>().ElementAt(8);
            Paragraph paragraph8 = body1.Elements<Paragraph>().ElementAt(10);
            Paragraph paragraph9 = body1.Elements<Paragraph>().ElementAt(11);
            Paragraph paragraph10 = body1.Elements<Paragraph>().ElementAt(13);
            SectionProperties sectionProperties1 = body1.GetFirstChild<SectionProperties>();

            Run run1 = paragraph1.GetFirstChild<Run>();

            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            paragraph1.InsertBefore(bookmarkStart1, run1);

            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };
            paragraph1.InsertBefore(bookmarkEnd1, run1);

            Run run2 = paragraph2.GetFirstChild<Run>();

            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph2.InsertBefore(proofError1, run2);

            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph2.Append(proofError2);

            Run run3 = paragraph3.GetFirstChild<Run>();

            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph3.InsertBefore(proofError3, run3);

            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph3.Append(proofError4);

            Run run4 = paragraph4.GetFirstChild<Run>();

            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph4.InsertBefore(proofError5, run4);

            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph4.Append(proofError6);

            Run run5 = paragraph5.GetFirstChild<Run>();

            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph5.InsertBefore(proofError7, run5);

            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph5.Append(proofError8);

            Run run6 = paragraph6.GetFirstChild<Run>();

            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph6.InsertBefore(proofError9, run6);

            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph6.Append(proofError10);

            Run run7 = paragraph7.GetFirstChild<Run>();

            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph7.InsertBefore(proofError11, run7);

            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph7.Append(proofError12);

            Run run8 = paragraph8.GetFirstChild<Run>();

            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            paragraph8.InsertBefore(proofError13, run8);

            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph8.Append(proofError14);

            Run run9 = paragraph9.GetFirstChild<Run>();

            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.GrammarStart };
            paragraph9.InsertBefore(proofError15, run9);

            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };
            paragraph9.Append(proofError16);

            BookmarkStart bookmarkStart2 = paragraph10.GetFirstChild<BookmarkStart>();
            BookmarkEnd bookmarkEnd2 = paragraph10.GetFirstChild<BookmarkEnd>();

            bookmarkStart2.Remove();
            bookmarkEnd2.Remove();

            PageSize pageSize1 = sectionProperties1.GetFirstChild<PageSize>();

            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Even, Id = "rId6" };
            sectionProperties1.InsertBefore(headerReference1, pageSize1);

            HeaderReference headerReference2 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId7" };
            sectionProperties1.InsertBefore(headerReference2, pageSize1);

            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Even, Id = "rId8" };
            sectionProperties1.InsertBefore(footerReference1, pageSize1);

            FooterReference footerReference2 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId9" };
            sectionProperties1.InsertBefore(footerReference2, pageSize1);

            HeaderReference headerReference3 = new HeaderReference() { Type = HeaderFooterValues.First, Id = "rId10" };
            sectionProperties1.InsertBefore(headerReference3, pageSize1);

            FooterReference footerReference3 = new FooterReference() { Type = HeaderFooterValues.First, Id = "rId11" };
            sectionProperties1.InsertBefore(footerReference3, pageSize1);
        }

        private void ChangeDocumentSettingsPart1(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = documentSettingsPart1.Settings;

            DefaultTabStop defaultTabStop1 = settings1.GetFirstChild<DefaultTabStop>();
            Compatibility compatibility1 = settings1.GetFirstChild<Compatibility>();
            Rsids rsids1 = settings1.GetFirstChild<Rsids>();
            ShapeDefaults shapeDefaults1 = settings1.GetFirstChild<ShapeDefaults>();

            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            settings1.InsertBefore(proofState1, defaultTabStop1);

            HeaderShapeDefaults headerShapeDefaults1 = new HeaderShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2052 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "2" };

            shapeLayout1.Append(shapeIdMap1);

            headerShapeDefaults1.Append(shapeDefaults2);
            headerShapeDefaults1.Append(shapeLayout1);
            settings1.InsertBefore(headerShapeDefaults1, compatibility1);

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);
            settings1.InsertBefore(footnoteDocumentWideProperties1, compatibility1);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);
            settings1.InsertBefore(endnoteDocumentWideProperties1, compatibility1);

            Rsid rsid1 = rsids1.GetFirstChild<Rsid>();

            Rsid rsid2 = new Rsid() { Val = "003F318F" };
            rsids1.InsertBefore(rsid2, rsid1);

            Rsid rsid3 = new Rsid() { Val = "00422D22" };
            rsids1.InsertBefore(rsid3, rsid1);

            Ovml.ShapeDefaults shapeDefaults3 = shapeDefaults1.GetFirstChild<Ovml.ShapeDefaults>();
            shapeDefaults3.MaxShapeId = 2052;
        }

        private void ChangeStyleDefinitionsPart1(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = styleDefinitionsPart1.Styles;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header" };
            StyleName styleName1 = new StyleName() { Val = "header" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "HeaderChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            Rsid rsid1 = new Rsid() { Val = "00422D22" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4320 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 8640 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(tabs1);
            styleParagraphProperties1.Append(spacingBetweenLines1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            styles1.Append(style1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "HeaderChar", CustomStyle = true };
            StyleName styleName2 = new StyleName() { Val = "Header Char" };
            BasedOn basedOn2 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Header" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            Rsid rsid2 = new Rsid() { Val = "00422D22" };

            style2.Append(styleName2);
            style2.Append(basedOn2);
            style2.Append(linkedStyle2);
            style2.Append(uIPriority2);
            style2.Append(rsid2);
            styles1.Append(style2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "Footer" };
            StyleName styleName3 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "FooterChar" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            Rsid rsid3 = new Rsid() { Val = "00422D22" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4320 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 8640 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties2.Append(tabs2);
            styleParagraphProperties2.Append(spacingBetweenLines2);

            style3.Append(styleName3);
            style3.Append(basedOn3);
            style3.Append(linkedStyle3);
            style3.Append(uIPriority3);
            style3.Append(unhideWhenUsed2);
            style3.Append(rsid3);
            style3.Append(styleParagraphProperties2);
            styles1.Append(style3);

            Style style4 = new Style() { Type = StyleValues.Character, StyleId = "FooterChar", CustomStyle = true };
            StyleName styleName4 = new StyleName() { Val = "Footer Char" };
            BasedOn basedOn4 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "Footer" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            Rsid rsid4 = new Rsid() { Val = "00422D22" };

            style4.Append(styleName4);
            style4.Append(basedOn4);
            style4.Append(linkedStyle4);
            style4.Append(uIPriority4);
            style4.Append(rsid4);
            styles1.Append(style4);
        }


    }
}
