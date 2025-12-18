using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Data.Pdf;         
using Windows.Storage.Streams; 

namespace SmartDocProcessor.WPF.Services
{
    public class OcrResultData
    {
        public string Text { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    public class OcrService
    {
        private OcrEngine? _ocrEngine;

        public OcrService() {
            TryInitOcr("ko-KR");
            if (_ocrEngine == null) TryInitOcr("en-US");
            if (_ocrEngine == null) _ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
        }

        private void TryInitOcr(string langCode) {
            if (_ocrEngine != null) return;
            try {
                var lang = new Language(langCode);
                if (OcrEngine.IsLanguageSupported(lang)) _ocrEngine = OcrEngine.TryCreateFromLanguage(lang);
            } catch { }
        }

        public string GetCurrentLanguage()
        {
            return _ocrEngine?.RecognizerLanguage.DisplayName ?? "OCR 엔진 없음";
        }

        // PDF 바이트를 받아 특정 페이지를 렌더링하고 OCR 수행
        public async Task PerformOcr(byte[] pdfBytes, List<AnnotationData> annotations, int pageIndex)
        {
            if (_ocrEngine == null || pdfBytes == null || pdfBytes.Length == 0) return;

            try
            {
                // 1. PDF 로드 (MemoryStream -> IRandomAccessStream)
                using (var randomAccessStream = new InMemoryRandomAccessStream())
                {
                    using (var writer = new DataWriter(randomAccessStream.GetOutputStreamAt(0))) {
                        writer.WriteBytes(pdfBytes);
                        await writer.StoreAsync();
                    }

                    var pdfDoc = await PdfDocument.LoadFromStreamAsync(randomAccessStream);
                    
                    if (pageIndex < 1 || pageIndex > pdfDoc.PageCount) return;

                    // 2. 해당 페이지 가져오기 (0-based index)
                    var pdfPage = pdfDoc.GetPage((uint)pageIndex - 1);

                    // 3. 이미지로 렌더링
                    using (var imageStream = new InMemoryRandomAccessStream())
                    {
                        await pdfPage.RenderToStreamAsync(imageStream);

                        // 4. OCR 수행
                        var decoder = await BitmapDecoder.CreateAsync(imageStream);
                        using (var softwareBitmap = await decoder.GetSoftwareBitmapAsync())
                        {
                            var result = await _ocrEngine.RecognizeAsync(softwareBitmap);

                            // 5. 결과를 AnnotationData로 변환하여 리스트에 추가
                            foreach (var line in result.Lines)
                            {
                                // 픽셀 좌표(96 DPI)를 PDF 포인트(72 DPI)로 대략 변환 (0.75 배)
                                double scaleFactor = 0.75; 
                                
                                double minX = double.MaxValue, minY = double.MaxValue;
                                double maxX = double.MinValue, maxY = double.MinValue;

                                foreach(var word in line.Words) {
                                    if(word.BoundingRect.X < minX) minX = word.BoundingRect.X;
                                    if(word.BoundingRect.Y < minY) minY = word.BoundingRect.Y;
                                    if(word.BoundingRect.Right > maxX) maxX = word.BoundingRect.Right;
                                    if(word.BoundingRect.Bottom > maxY) maxY = word.BoundingRect.Bottom;
                                }

                                var ann = new AnnotationData
                                {
                                    Type = "OCR_TEXT",
                                    Content = line.Text,
                                    X = minX * scaleFactor,
                                    Y = minY * scaleFactor,
                                    Width = (maxX - minX) * scaleFactor,
                                    Height = (maxY - minY) * scaleFactor,
                                    Page = pageIndex,
                                    Color = "#000000",
                                    FontSize = 10 
                                };
                                annotations.Add(ann);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PerformOcr Error: {ex.Message}");
            }
        }
    }
}