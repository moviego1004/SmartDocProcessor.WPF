using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Reflection; 
using System.Text.RegularExpressions;
using System.Globalization; 
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Fonts; 
using PdfSharp.Pdf.Annotations; 
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace SmartDocProcessor.WPF.Services
{
    public class AnnotationData
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Type { get; set; } = "TEXT";
        public string Content { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public int Page { get; set; }
        public string Color { get; set; } = "#000000";
        public int FontSize { get; set; } = 14; 
        public bool Selected { get; set; } = false;
        public string FontFamily { get; set; } = "Malgun Gothic";
        public bool IsBold { get; set; } = false;
    }

    public class PdfService
    {
        public PdfService() {
            try { 
                if (GlobalFontSettings.FontResolver == null) 
                    GlobalFontSettings.FontResolver = new SystemFontResolver(); 
            } catch { }
        }

        public bool IsPdfSearchable(byte[] pdfBytes)
        {
            try {
                using (var inStream = new MemoryStream(pdfBytes)) {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Import);
                    int pagesToCheck = Math.Min(doc.PageCount, 3);
                    for (int i = 0; i < pagesToCheck; i++) {
                        var content = ContentReader.ReadContent(doc.Pages[i]);
                        if (HasTextOperators(content)) return true;
                    }
                }
            } catch { }
            return false;
        }

        private bool HasTextOperators(CObject content)
        {
            if (content is COperator op) {
                if (op.OpCode.Name == "Tj" || op.OpCode.Name == "TJ" || op.OpCode.Name == "\'") return true;
            } else if (content is CSequence seq) {
                foreach (var item in seq) if (HasTextOperators(item)) return true;
            }
            return false;
        }

        public byte[] GetPdfBytesWithoutAnnotations(byte[] pdfBytes)
        {
            try {
                using (var inStream = new MemoryStream(pdfBytes))
                using (var outStream = new MemoryStream()) {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Modify);
                    foreach (var page in doc.Pages) if (page.Annotations != null) page.Annotations.Clear();
                    doc.Save(outStream);
                    return outStream.ToArray();
                }
            } catch { return pdfBytes; }
        }

        public List<AnnotationData> ExtractAnnotationsFromMetadata(byte[] pdfBytes)
        {
            var result = new List<AnnotationData>();
            try {
                using (var inStream = new MemoryStream(pdfBytes)) {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Import);
                    for (int i = 0; i < doc.PageCount; i++) {
                        var page = doc.Pages[i];
                        int pageNum = i + 1;
                        
                        // [수정] 좌표 계산: CropBox Top(Y2) 기준
                        // PdfSharp 6.x에서 Height, Y1 등은 이미 double (Point)
                        double pageTop = page.CropBox.Height > 0 ? (page.CropBox.Y1 + page.CropBox.Height) : page.Height;
                        double xOffset = page.CropBox.X1;

                        if (page.Annotations != null) {
                            foreach (var item in page.Annotations) {
                                if (item is PdfAnnotation annot) {
                                    var subType = annot.Elements.GetString("/Subtype");
                                    var rect = annot.Rectangle;
                                    
                                    // PDF(Bottom-Up) -> WPF(Top-Down) 변환
                                    // WPF Y = PageTop - Rect.Top(Y2)
                                    double appY = pageTop - rect.Y2;
                                    double appX = rect.X1 - xOffset;

                                    var data = new AnnotationData { X=appX, Y=appY, Width=rect.Width, Height=rect.Height, Page=pageNum, Content=annot.Contents ?? "" };
                                    
                                    if (subType == "/Highlight") { data.Type = "HIGHLIGHT_Y"; data.Color = "#FFFF00"; result.Add(data); }
                                    else if (subType == "/Underline") { data.Type = "UNDERLINE"; data.Color = "#FF0000"; result.Add(data); }
                                    else if (subType == "/FreeText") { 
                                        data.Type = "TEXT"; 
                                        ParseAppearanceString(annot.Elements.GetString("/DA"), data); 
                                        result.Add(data); 
                                    }
                                }
                            }
                        }
                    }
                }
            } catch { }
            return result;
        }

        private void ParseAppearanceString(string da, AnnotationData data)
        {
            if (string.IsNullOrEmpty(da)) return;
            try {
                var fontMatch = Regex.Match(da, @"([\d\.]+)\s+Tf");
                if (fontMatch.Success && double.TryParse(fontMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double size)) {
                    // PDF(72 DPI) -> WPF(96 DPI) 복원
                    data.FontSize = (int)Math.Round(size * 96.0 / 72.0);
                }
                
                var rgbMatch = Regex.Match(da, @"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+rg");
                if (rgbMatch.Success) {
                    double r = double.Parse(rgbMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    double g = double.Parse(rgbMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                    double b = double.Parse(rgbMatch.Groups[3].Value, CultureInfo.InvariantCulture);
                    data.Color = $"#FF{(int)(r*255):X2}{(int)(g*255):X2}{(int)(b*255):X2}";
                }
            } catch { }
        }

        public byte[] SavePdfWithAnnotations(byte[] originalPdf, List<AnnotationData> annotations, double currentScale)
        {
            using (var inStream = new MemoryStream(originalPdf))
            using (var outStream = new MemoryStream())
            {
                var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Modify);
                try { 
                    if (!doc.AcroForm.Elements.ContainsKey("/NeedAppearances")) doc.AcroForm.Elements.Add("/NeedAppearances", new PdfBoolean(true)); 
                    else doc.AcroForm.Elements["/NeedAppearances"] = new PdfBoolean(true); 
                } catch { }

                foreach (var page in doc.Pages) if(page.Annotations != null) page.Annotations.Clear();

                foreach (var group in annotations.GroupBy(a => a.Page))
                {
                    int pageIdx = group.Key - 1;
                    if (pageIdx < 0 || pageIdx >= doc.PageCount) continue;
                    var page = doc.Pages[pageIdx];
                    
                    // [중요] 좌표 계산의 기준점: Page Top
                    // CropBox가 있으면 그 높이+Y1을 Top으로, 없으면 Page Height를 Top으로
                    double pageTop = page.CropBox.Height > 0 ? (page.CropBox.Y1 + page.CropBox.Height) : page.Height;
                    double xOffset = page.CropBox.X1;

                    // 1. OCR 텍스트 (투명 저장)
                    var ocrItems = group.Where(a => a.Type == "OCR_TEXT");
                    if (ocrItems.Any())
                    {
                        using (var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append))
                        {
                            foreach (var ann in ocrItems)
                            {
                                var brush = new XSolidBrush(XColor.FromArgb(1, 0, 0, 0)); // 거의 완전 투명
                                var font = new XFont("Arial", ann.FontSize, XFontStyleEx.Regular);
                                // OCR 좌표는 그대로 사용 (이미지 기준) -> XGraphics도 TopLeft 기준이므로 OK
                                gfx.DrawString(ann.Content, font, brush, new XRect(ann.X, ann.Y, ann.Width, ann.Height), XStringFormats.TopLeft);
                            }
                        }
                    }

                    // 2. 사용자 주석
                    foreach (var ann in group.Where(a => a.Type != "OCR_TEXT"))
                    {
                        // [핵심] 좌표 변환: WPF(Top-Down) -> PDF(Bottom-Up)
                        // PDF Rect Top (Y2) = PageTop - WPF_Y
                        // PDF Rect Bottom (Y1) = PDF_Top - Height
                        double pdfTopY = pageTop - ann.Y;
                        double pdfBottomY = pdfTopY - ann.Height;
                        double pdfLeftX = xOffset + ann.X;

                        // [수정] Point 단위 사용
                        var rect = new PdfRectangle(new XPoint(pdfLeftX, pdfBottomY), new XPoint(pdfLeftX + ann.Width, pdfTopY));

                        if (ann.Type == "TEXT")
                        {
                            var annot = new CustomPdfAnnotation(doc);
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, "/FreeText");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            annot.Elements.SetString(PdfAnnotation.Keys.Contents, ann.Content);

                            XColor c = XColor.FromArgb(255, 0, 0, 0);
                            try {
                                string hex = ann.Color.Replace("#", "");
                                if (hex.Length == 8) c = XColor.FromArgb(int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber), int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber), int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber), int.Parse(hex.Substring(6, 2), NumberStyles.HexNumber));
                                else if (hex.Length == 6) c = XColor.FromArgb(255, int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber), int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber), int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber));
                            } catch {}

                            double r = c.R / 255.0; double g = c.G / 255.0; double b = c.B / 255.0;
                            
                            // [보정] 폰트 크기: WPF(96) -> PDF(72) = 0.75배
                            double pdfFontSize = ann.FontSize * 0.75;

                            annot.Elements["/DA"] = new PdfString($"/Helv {pdfFontSize:0.##} Tf {r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b.ToString("0.###", CultureInfo.InvariantCulture)} rg");

                            var formRect = new XRect(0, 0, ann.Width, ann.Height);
                            var form = new XForm(doc, formRect);
                            
                            using (var gfx = XGraphics.FromForm(form))
                            {
                                var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                var font = CreateSafeFont(ann.FontFamily, Math.Max(pdfFontSize, 4), ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular, options);
                                var brush = new XSolidBrush(c);
                                var tf = new XTextFormatter(gfx);
                                tf.Alignment = XParagraphAlignment.Left;
                                // WPF TextBox Padding=0에 맞춰 (0,0)에 그리기
                                tf.DrawString(ann.Content, font, brush, formRect, XStringFormats.TopLeft);
                            }
                            var apDict = new PdfDictionary(doc);
                            annot.Elements["/AP"] = apDict;
                            var pdfForm = GetPdfForm(form);
                            if (pdfForm != null) apDict.Elements["/N"] = pdfForm;
                            page.Annotations.Add(annot);
                        }
                        else 
                        {
                            var annot = new CustomPdfAnnotation(doc);
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, ann.Type.StartsWith("HIGHLIGHT") ? "/Highlight" : "/Underline");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            
                            if(ann.Type.StartsWith("HIGHLIGHT")) {
                                var color = ann.Type == "HIGHLIGHT_O" ? XColor.FromArgb(255, 165, 0) : XColor.FromArgb(255, 255, 0);
                                annot.Elements["/C"] = new PdfArray(doc, new PdfReal(color.R / 255.0), new PdfReal(color.G / 255.0), new PdfReal(color.B / 255.0));
                            } else {
                                annot.Elements["/C"] = new PdfArray(doc, new PdfReal(1), new PdfReal(0), new PdfReal(0));
                            }

                            // [중요] QuadPoints (Z 순서: 좌상, 우상, 좌하, 우하) - PDF 좌표계 기준
                            // PDF 좌표계에서는 Y가 클수록 위쪽임.
                            // Point 1: TL (Left, Top) -> (rect.X1, rect.Y2)
                            // Point 2: TR (Right, Top) -> (rect.X2, rect.Y2)
                            // Point 3: BL (Left, Bottom) -> (rect.X1, rect.Y1)
                            // Point 4: BR (Right, Bottom) -> (rect.X2, rect.Y1)
                            
                            var qp = new PdfArray(doc);
                            qp.Elements.Add(new PdfReal(rect.X1)); qp.Elements.Add(new PdfReal(rect.Y2)); // TL
                            qp.Elements.Add(new PdfReal(rect.X2)); qp.Elements.Add(new PdfReal(rect.Y2)); // TR
                            qp.Elements.Add(new PdfReal(rect.X1)); qp.Elements.Add(new PdfReal(rect.Y1)); // BL
                            qp.Elements.Add(new PdfReal(rect.X2)); qp.Elements.Add(new PdfReal(rect.Y1)); // BR
                            
                            annot.Elements["/QuadPoints"] = qp;
                            page.Annotations.Add(annot);
                        }
                    }
                }
                doc.Save(outStream);
                return outStream.ToArray();
            }
        }

        private XFont CreateSafeFont(string familyName, double size, XFontStyleEx style, XPdfFontOptions? options = null)
        {
            if (options == null) options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            try { return new XFont(familyName, size, style, options); }
            catch { try { return new XFont("Malgun Gothic", size, style, options); } catch { return new XFont("Arial", size, style, options); } }
        }

        public byte[] DeletePage(byte[] pdfBytes, int pageIndex) { using (var ms = new MemoryStream(pdfBytes)) using (var outMs = new MemoryStream()) { var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import); var newDoc = new PdfDocument(); for (int i = 0; i < doc.PageCount; i++) if (i != pageIndex) newDoc.AddPage(doc.Pages[i]); newDoc.Save(outMs); return outMs.ToArray(); } }

        private PdfFormXObject? GetPdfForm(XForm form)
        {
            try { var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public); if (prop != null) return prop.GetValue(form) as PdfFormXObject; } catch { }
            return null;
        }
    }

    public class CustomPdfAnnotation : PdfAnnotation
    {
        public CustomPdfAnnotation(PdfDocument document) : base(document) { }
    }
}