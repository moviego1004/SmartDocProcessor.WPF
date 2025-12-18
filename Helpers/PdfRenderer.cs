using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace SmartDocProcessor.WPF.Helpers
{
    public static class PdfRenderer
    {
        // PDF 바이트 배열에서 특정 페이지를 WPF 이미지로 변환
        public static async Task<BitmapSource> RenderPageToBitmapAsync(byte[] pdfData, int pageIndex)
        {
            using (var stream = new InMemoryRandomAccessStream())
            {
                using (var writer = new DataWriter(stream.GetOutputStreamAt(0)))
                {
                    writer.WriteBytes(pdfData);
                    await writer.StoreAsync();
                }

                var pdfDoc = await PdfDocument.LoadFromStreamAsync(stream);
                if (pageIndex < 1 || pageIndex > pdfDoc.PageCount) return null;

                var page = pdfDoc.GetPage((uint)pageIndex - 1);

                // 고해상도 렌더링을 위해 스케일 조정 가능 (기본 1.0)
                var renderOptions = new PdfPageRenderOptions { DestinationWidth = (uint)(page.Size.Width * 1.5) }; 
                
                using (var imageStream = new InMemoryRandomAccessStream())
                {
                    await page.RenderToStreamAsync(imageStream, renderOptions);
                    
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = imageStream.AsStreamForRead();
                    bitmap.EndInit();
                    bitmap.Freeze(); // UI 스레드 간 공유를 위해 Freeze 필수
                    return bitmap;
                }
            }
        }
    }
}