using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GhostScriptNS
{
    class Program
    {
        static void Main(string[] args)
        {

            //MergePdf();

            //Get All Files ina  directory
            List<FileInfo> _Files = new List<FileInfo>();

            foreach (var file in new DirectoryInfo(@"D:\PDF\Input").GetFiles("*.pdf", SearchOption.AllDirectories).OrderBy(x=>x.Name))
            {
                _Files.Add(file);
            }

            Parallel.ForEach(_Files, file =>
            {
                PdfReader reader = new PdfReader(file.FullName);

                //Compress images, Specify quality as second paramter - 1 to 100
                ReduceResolution(reader, 7);

                if (!Directory.Exists($@"D:\PDF\Output\{file.Directory.Name}")) Directory.CreateDirectory($@"D:\PDF\Output\{file.Directory.Name}");

                using (PdfStamper stamper = new PdfStamper(reader, new FileStream($@"D:\PDF\Output\{file.Directory.Name}\{file.Name}", FileMode.Create), PdfWriter.VERSION_1_7))
                {
                    // flatten form fields and close document
                    stamper.FormFlattening = true;

                    stamper.SetFullCompression();

                    //IText document - http://itextsupport.com/apidocs/itext5/latest/com/itextpdf/text/pdf/PdfStamper.html
                    //strength - true for 128 bit key length, false for 40 bit key length
                    stamper.SetEncryption(true, "tk", "tk", PdfWriter.AllowPrinting);
                    stamper.Close();
                }


                reader.Close();
            });

            //MergePdf();
            _Files.Clear();
            _Files = null;
        }

        /// <summary>
        /// Creates a single PDF from multiple at a specified location        
        /// </summary>
        private static void MergePdf()
        {
            String result = @"d:\PDF\Output\mergefile.pdf";
            
            Document document = new Document();
            //create PdfSmartCopy object
            PdfSmartCopy copy = new PdfSmartCopy(document, new FileStream(result, FileMode.Create));
            copy.SetFullCompression();
            //open the document
            document.Open();
            
            PdfReader reader = null;
            foreach (var file in new DirectoryInfo(@"D:\PDF\Output").GetFiles("*.pdf").OrderBy(x => x.Name))
            {
                if (file.Name == "mergefile.pdf") continue;
                //create PdfReader object
                reader = new PdfReader(file.FullName);
                //merge combine pages
                for (int page = 1; page <= reader.NumberOfPages; page++)
                    copy.AddPage(copy.GetImportedPage(reader, page));

            }
            //close the document object
            document.Close();
        }


        /// <summary>
        /// Gets image from PDF and compresses it - Found on StackOverflow - asis
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="quality"></param>
        public static void ReduceResolution(PdfReader reader, long quality)
        {
            int n = reader.XrefSize;
            for (int i = 0; i < n; i++)
            {
                PdfObject obj = reader.GetPdfObject(i);
                if (obj == null || !obj.IsStream()) { continue; }

                PdfDictionary dict = (PdfDictionary)PdfReader.GetPdfObject(obj);
                PdfName subType = (PdfName)PdfReader.GetPdfObject(
                  dict.Get(PdfName.SUBTYPE)
                );
                if (!PdfName.IMAGE.Equals(subType)) { continue; }

                PRStream stream = (PRStream)obj;
                try
                {
                    PdfImageObject image = new PdfImageObject(stream);
                    //PdfName filter = (PdfName)image.Get(PdfName.FILTER);
                    //if (
                    //  PdfName.JBIG2DECODE.Equals(filter)
                    //  || PdfName.JPXDECODE.Equals(filter)
                    //  || PdfName.CCITTFAXDECODE.Equals(filter)
                    //  || PdfName.FLATEDECODE.Equals(filter)
                    //) continue;

                    System.Drawing.Image img = image.GetDrawingImage();
                    if (img == null) continue;

                    var ll = image.GetImageBytesType();
                    int width = img.Width;
                    int height = img.Height;
                    using (System.Drawing.Bitmap dotnetImg =
                       new System.Drawing.Bitmap(img))
                    {
                        // set codec to jpeg type => jpeg index codec is "1"
                        System.Drawing.Imaging.ImageCodecInfo codec =
                        System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()[1];
                        // set parameters for image quality
                        System.Drawing.Imaging.EncoderParameters eParams =
                         new System.Drawing.Imaging.EncoderParameters(1);
                        eParams.Param[0] =
                         new System.Drawing.Imaging.EncoderParameter(
                           System.Drawing.Imaging.Encoder.Quality, quality
                        );
                        using (MemoryStream msImg = new MemoryStream())
                        {
                            dotnetImg.Save(msImg, codec, eParams);
                            msImg.Position = 0;
                            stream.SetData(msImg.ToArray());
                            stream.SetData(
                             msImg.ToArray(), false, PRStream.BEST_COMPRESSION
                            );
                            stream.Put(PdfName.TYPE, PdfName.XOBJECT);
                            stream.Put(PdfName.SUBTYPE, PdfName.IMAGE);
                            stream.Put(PdfName.FILTER, image.Get(PdfName.FILTER));
                            stream.Put(PdfName.FILTER, PdfName.DCTDECODE);
                            stream.Put(PdfName.WIDTH, new PdfNumber(width));
                            stream.Put(PdfName.HEIGHT, new PdfNumber(height));
                            stream.Put(PdfName.BITSPERCOMPONENT, new PdfNumber(8));
                            stream.Put(PdfName.COLORSPACE, PdfName.DEVICERGB);
                        }
                    }
                }
                catch
                {
                    // throw;
                    // iText[Sharp] can't handle all image types...
                }
                finally
                {
                    // may or may not help      
                    reader.RemoveUnusedObjects();
                }
            }


        }
    }
}
