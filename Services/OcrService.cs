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
    // [신규] 텍스트 위치 정보를 담을 클래스
    public class TextData
    {
        public string Text { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

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

        public string GetCurrentLanguage() => _ocrEngine?.RecognizerLanguage.DisplayName ?? "OCR 엔진 없음";

        // [신규] 텍스트 데이터 추출 (드래그 선택용)
        public async Task<List<TextData>> ExtractTextData(byte[] pdfData, int pageIndex)
        {
            var list = new List<TextData>();
            if (_ocrEngine == null || pdfData == null) return list;

            try
            {
                using (var stream = new InMemoryRandomAccessStream())
                {
                    using (var writer = new DataWriter(stream.GetOutputStreamAt(0))) {
                        writer.WriteBytes(pdfData);
                        await writer.StoreAsync();
                    }

                    var pdfDoc = await PdfDocument.LoadFromStreamAsync(stream);
                    if (pageIndex < 1 || pageIndex > pdfDoc.PageCount) return list;
                    var page = pdfDoc.GetPage((uint)pageIndex - 1);

                    // OCR 좌표 정밀도를 위해 2배 확대
                    double scale = 2.0;
                    var options = new PdfPageRenderOptions { DestinationWidth = (uint)(page.Size.Width * scale) };

                    using (var imgStream = new InMemoryRandomAccessStream())
                    {
                        await page.RenderToStreamAsync(imgStream, options);
                        var decoder = await BitmapDecoder.CreateAsync(imgStream);
                        using (var softwareBitmap = await decoder.GetSoftwareBitmapAsync())
                        {
                            var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);
                            foreach (var line in ocrResult.Lines)
                            {
                                foreach (var word in line.Words)
                                {
                                    var rect = word.BoundingRect;
                                    list.Add(new TextData
                                    {
                                        Text = word.Text,
                                        X = rect.X / scale,
                                        Y = rect.Y / scale,
                                        Width = rect.Width / scale,
                                        Height = rect.Height / scale
                                    });
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        public async Task PerformOcr(byte[] pdfBytes, List<AnnotationData> annotations, int pageIndex)
        {
            var texts = await ExtractTextData(pdfBytes, pageIndex);
            foreach (var t in texts)
            {
                annotations.Add(new AnnotationData
                {
                    Type = "OCR_TEXT",
                    Content = t.Text,
                    X = t.X, Y = t.Y, Width = t.Width, Height = t.Height,
                    Page = pageIndex,
                    Color = "#000000", FontSize = 10 
                });
            }
        }
    }
}