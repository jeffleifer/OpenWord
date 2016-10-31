using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Font = System.Drawing.Font;
using FontFamily = System.Drawing.FontFamily;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
namespace WordOpen.Templates
{
    internal class ProcessImage : BaseProcess
    {

        
        internal static void ShowIndicator(WordprocessingDocument docx,
           
            SimpleField field,
            string measured, 
            float operatorValue, 
            float inspectorValue, 
            float sigma)
        {
            var paragraph  = GetFirstParent<Paragraph>(field);
           
          
            MainDocumentPart mainPart = docx.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            var fileName = CreateImage(measured, operatorValue,inspectorValue,sigma);
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToBody(docx, mainPart.GetIdOfPart(imagePart),paragraph);

        }
        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, Paragraph paragraph)
        {
            // Define the reference of the image.
            var element =
          new Drawing(
              new DW.Inline(
                  new DW.Extent() { Cx = 990000L, Cy = 792000L },
                  new DW.EffectExtent()
                  {
                      LeftEdge = 0L,
                      TopEdge = 0L,
                      RightEdge = 0L,
                      BottomEdge = 0L
                  },
                  new DW.DocProperties()
                  {
                      Id = (UInt32Value)1U,
                      Name = "Picture 1"
                  },
                  new DW.NonVisualGraphicFrameDrawingProperties(
                      new A.GraphicFrameLocks() { NoChangeAspect = true }),
                  new A.Graphic(
                      new A.GraphicData(
                          new PIC.Picture(
                              new PIC.NonVisualPictureProperties(
                                  new PIC.NonVisualDrawingProperties()
                                  {
                                      Id = (UInt32Value)0U,
                                      Name = "New Bitmap Image.jpg"
                                  },
                                  new PIC.NonVisualPictureDrawingProperties()),
                              new PIC.BlipFill(
                                  new A.Blip(
                                      new A.BlipExtensionList(
                                          new A.BlipExtension()
                                          {
                                              Uri =
                                                 "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                          })
                                  )
                                  {
                                      Embed = relationshipId,
                                      CompressionState =
                                      A.BlipCompressionValues.Print
                                  },
                                  new A.Stretch(
                                      new A.FillRectangle())),
                              new PIC.ShapeProperties(
                                  new A.Transform2D(
                                      new A.Offset() { X = 0L, Y = 0L },
                                      new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                  new A.PresetGeometry(
                                      new A.AdjustValueList()
                                  )
                                  { Preset = A.ShapeTypeValues.Rectangle }))
                      )
                      { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
              )
              {
                  DistanceFromTop = (UInt32Value)0U,
                  DistanceFromBottom = (UInt32Value)0U,
                  DistanceFromLeft = (UInt32Value)0U,
                  DistanceFromRight = (UInt32Value)0U,
                  EditId = "50D07946"
              });

            // Append the reference to body, the element should be in a Run.
            paragraph.RemoveAllChildren();
            paragraph.AppendChild(new Run(element));
           // wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

      

        private static string CreateImage(string measured, float operatorValue, float inspectorValue, float sigma)
        {
            string path = @"C:\Users\X230\Documents\DocGen\test.png";// + DateTime.Now.ToString("HHmmss") + ".png";
            var templateImage = @"C:\Users\X230\Documents\DocGen\template.png";
            var templateArrowImage = @"C:\Users\X230\Documents\DocGen\templateArrow.png";
            if (File.Exists(path))
                File.Delete(path);
           
            
            Font font = new Font(FontFamily.GenericSerif, 10);
            using (var img = Image.FromFile(templateImage))
            {
                var value = operatorValue - inspectorValue;
                var stringValue = value.ToString("##.000");

                var minValue = (float)-7.5*sigma;
                var maxValue = -1*minValue;
                var width = img.Width;
                var x = (float)(width*((value - minValue)/(2*maxValue)));
               
                using (var g = Graphics.FromImage(img))
                {
                    var arrowImage = Image.FromFile(templateArrowImage);
                    if (x > width)
                        x = width - arrowImage.Width/2;
                    else if (x < 0)
                        x = arrowImage.Width/2;
                    g.DrawImage(arrowImage, new PointF(x-arrowImage.Width/2, 0));
                    g.DrawString(measured, font, Brushes.Black, x - arrowImage.Width/6, arrowImage.Height/4);
                    var size = g.MeasureString(stringValue, font);
                    
                    var r = new RectangleF( x - size.Width / 2, arrowImage.Height , size.Width, size.Height);
                    var brush = (Math.Abs(value) <= 3 * sigma ? Brushes.Green : Brushes.DarkRed);
                    g.FillRectangle(brush, r);
                    g.DrawString(stringValue, font, Brushes.Black, r.X , r.Y + r.Height / 6);
                   


                }
                img.Save(path, ImageFormat.Png);
            }
            return path;


        }
    }
}