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
using Microsoft.SharePoint;
using System.Xml;

namespace shrenky.projects.watermark.console
{
    class Program
    {
        static void Main(string[] args)
        {
            UploadFileToSP();
            return;
            GetItemXmlPropertiesTest();
            return;
            using (WordprocessingDocument package = WordprocessingDocument.Open(@"C:\Users\Administrator\Desktop\SamTest.docx", true))
            {
                WatermarkHandler.InsertCustomWatermark(package, @"C:\Users\Administrator\Desktop\jeff.jpg");
            }
            Console.ReadLine();
            return;
            string filePath = @"C:\Users\Administrator\Desktop\Test.docx";
            string newFilePath = @"C:\Users\Administrator\Desktop\Test_New.docx";

            byte[] sourceBytes = File.ReadAllBytes(filePath);
            MemoryStream inMemoryStream = new MemoryStream();
            inMemoryStream.Write(sourceBytes, 0, (int)sourceBytes.Length);

            WordprocessingDocument doc = WordprocessingDocument.Open(inMemoryStream, true);
            WatermarkHandler.AddText(doc, "Test");
            WatermarkHandler.AddHeader(doc, "header");

            using (FileStream fileStream = new FileStream(newFilePath, System.IO.FileMode.Create))
            {
                inMemoryStream.WriteTo(fileStream);
            }

            inMemoryStream.Close();
            inMemoryStream.Dispose();
            inMemoryStream = null;
        }

        private static void UploadFileToSP()
        {
            string filePath = @"C:\Users\Administrator\Desktop\error.png";

            byte[] content = File.ReadAllBytes(filePath);
            string fileName = "TestPicture.png";
            using (SPSite site = new SPSite("http://server2013"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList pictureStore = web.Lists["Pictures"];
                    string destUrl = string.Format("/{0}/{1}", pictureStore.RootFolder.Url, fileName);
                    pictureStore.RootFolder.Files.Add(destUrl, content);
                    pictureStore.RootFolder.Update();
                }
            }
        }

        public Stream FileToStream(string fileName)
        {
            // 打开文件 
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            // 读取文件的 byte[] 
            byte[] bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, bytes.Length);
            fileStream.Close();
            // 把 byte[] 转换成 Stream 
            Stream stream = new MemoryStream(bytes);
            return stream;
        }
        private static void GetItemXmlPropertiesTest()
        {
            using (SPSite site = new SPSite("http://server2013"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList lib = web.Lists["Pictures"];
                    SPListItem item = lib.GetItemById(1);
                    string xml = item.Xml;
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xml);
                    /*ows_ThumbnailExists='1' 
                     * ows_PreviewExists='1' 
                     * ows_EncodedAbsThumbnailUrl='http://dev2013/Pictures/_t/Snap1_bmp.jpg' 
                     * ows_EncodedAbsWebImgUrl='http://dev2013/Pictures/_w/Snap1_bmp.jpg'
                     */
                    Console.WriteLine(doc.FirstChild.Attributes["ows_ThumbnailExists"].Value);
                    Console.WriteLine(doc.FirstChild.Attributes["ows_PreviewExists"].Value);
                    Console.WriteLine(doc.FirstChild.Attributes["ows_EncodedAbsThumbnailUrl"].Value);
                    Console.WriteLine(doc.FirstChild.Attributes["ows_EncodedAbsWebImgUrl"].Value);
                }
            }
        }

        static Header MakeHeader()
        {
            var header = new Header();
            var paragraph = new Paragraph();
            var run = new Run();
            var text = new Text();
            text.Text = "";
            run.Append(text);
            paragraph.Append(run);
            header.Append(paragraph);
            return header;
        }

       
    }

    public static class WatermarkHandler
    {
        public static bool AddText(WordprocessingDocument doc, string text)
        {
            bool result = true;
            try
            {
                using (doc)
                {
                    Document document = doc.MainDocumentPart.Document;
                    Paragraph firstParagraph = document.Body.Elements<Paragraph>().FirstOrDefault();
                    if (firstParagraph != null)
                    {
                        Paragraph testParagraph = new Paragraph(
                            new Run(
                                new Text(text)));
                        firstParagraph.Parent.InsertBefore(testParagraph,
                            firstParagraph);
                    }
                }
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public static bool AddHeader(WordprocessingDocument doc, string text)
        {
            bool result = true;
            try
            {

            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public static void InsertCustomWatermark(WordprocessingDocument package, string imagePath)
        {
            string imagePart1Data = SetWatermarkPicture(imagePath);
            MainDocumentPart mainDocumentPart1 = package.MainDocumentPart;
            if (mainDocumentPart1 != null)
            {
                mainDocumentPart1.DeleteParts(mainDocumentPart1.HeaderParts);
                HeaderPart headPart1 = mainDocumentPart1.AddNewPart<HeaderPart>();
                GenerateHeaderPart1Content(headPart1);
                string rId = mainDocumentPart1.GetIdOfPart(headPart1);
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/jpeg", "rId999");
                GenerateImagePart1Content(image, imagePart1Data);
                IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
                }
            }
            else
            {
                Console.WriteLine("alert");
            }
        }

        #region private
        private static void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();
            Paragraph paragraph2 = new Paragraph();
            Run run1 = new Run();
            Picture picture1 = new Picture();
            V.Shape shape1 = new V.Shape() { Id = "WordPictureWatermark75517470", Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:415.2pt;height:456.15pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2051", AllowInCell = false, Type = "#_x0000_t75" };
            V.ImageData imageData1 = new V.ImageData() { Gain = "19661f", BlackLevel = "22938f", Title = "水印", RelationshipId = "rId999" };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph2.Append(run1);
            header1.Append(paragraph2);
            headerPart1.Header = header1;
        }

        private static void GenerateImagePart1Content(ImagePart imagePart1, string imagePart1Data)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private static System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private static string SetWatermarkPicture(string path)
        {
            FileStream inFile;
            byte[] byteArray;
            string imagePart1Data = string.Empty;
            try
            {
                inFile = new FileStream(path, FileMode.Open, FileAccess.Read);
                byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int)inFile.Length);
                inFile.Close();
                imagePart1Data = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return imagePart1Data;
        }
        #endregion
    }
}
