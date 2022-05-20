using System;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.IO;
using System.Drawing;
using System.Threading.Tasks;

namespace AsposePOC
{
    class Program
    {
        private void SaveSlideToImage(ISlide sld)
        {
            Bitmap bmp = sld.GetThumbnail(1f, 1f);

            // Save the image to disk in JPEG format
            bmp.Save(string.Format(@"C:\Users\60102\source\repos\AsposePOC\AsposePOC\Images\Slide_{0}.jpg", sld.SlideNumber), System.Drawing.Imaging.ImageFormat.Jpeg);
        }
        static void Main(string[] args)
        {
            const string pathToVeryLargePresentationFile = @"C:\Users\60102\source\repos\AsposePOC\AsposePOC\pptx\Test.pptx";

            LoadOptions loadOptions = new LoadOptions
            {
                BlobManagementOptions = {
                    PresentationLockingBehavior = PresentationLockingBehavior.KeepLocked,
                 }
            };

            using (Presentation pres = new Presentation(pathToVeryLargePresentationFile))
            {
                //OPTION 1
                
                //foreach (ISlide sld in pres.Slides)
                //{
                //    // Create a full scale image
                //    Bitmap bmp = sld.GetThumbnail(1f, 1f);

                //    // Save the image to disk in JPEG format
                //    bmp.Save(string.Format(@"C:\Users\60102\source\repos\AsposePOC\AsposePOC\Images\Slide_{0}.jpg", sld.SlideNumber), System.Drawing.Imaging.ImageFormat.Jpeg);
                //}

                //OPTION 2 USING PTL
                
                Parallel.ForEach(pres.Slides, slide => new Program().SaveSlideToImage(slide));


            }


        }        
    }
}
