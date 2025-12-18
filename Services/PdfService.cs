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

        // PDF가 Searchable(텍스트 포함)인지 확인
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
                        
                        // [수정] 좌표 계산 시 CropBox 오프셋 고려
                        double pageHeight = page.CropBox.Height > 0 ? page.CropBox.Height : page.Height;
                        double yOffset = page.CropBox.Y1; // CropBox 시작 Y (Bottom)
                        double xOffset = page.CropBox.X1; // CropBox 시작 X (Left)

                        if (page.Annotations != null) {
                            foreach (var item in page.Annotations) {
                                if (item is PdfAnnotation annot) {
                                    var subType = annot.Elements.GetString("/Subtype");
                                    var rect = annot.Rectangle;
                                    
                                    // PDF(Bottom-Up) -> WPF(Top-Down) 변환
                                    // VisualTop = (PageHeight + OffsetY) - Rect.Top
                                    double appY = (pageHeight + yOffset) - rect.Y2;
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
                    
                    // [중요] CropBox 기준 좌표 정보 가져오기
                    // PDFSharp의 page.Height는 MediaBox 기준일 수 있으므로 CropBox를 우선함
                    double pageHeight = page.CropBox.Height > 0 ? page.CropBox.Height : page.Height;
                    double xOffset = page.CropBox.X1;
                    double yOffset = page.CropBox.Y1;

                    // 1. OCR 텍스트 (투명)
                    var ocrItems = group.Where(a => a.Type == "OCR_TEXT");
                    if (ocrItems.Any())
                    {
                        using (var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append))
                        {
                            foreach (var ann in ocrItems)
                            {
                                var brush = new XSolidBrush(XColor.FromArgb(1, 0, 0, 0)); // 완전 투명
                                var font = new XFont("Arial", ann.FontSize, XFontStyleEx.Regular);
                                // OCR 좌표는 이미 PDF 좌표계와 일치한다고 가정하고 그대로 사용
                                gfx.DrawString(ann.Content, font, brush, new XRect(ann.X, ann.Y, ann.Width, ann.Height), XStringFormats.TopLeft);
                            }
                        }
                    }

                    // 2. 사용자 주석
                    foreach (var ann in group.Where(a => a.Type != "OCR_TEXT"))
                    {
                        // [핵심] 좌표 변환: WPF(Top-Down) -> PDF(Bottom-Up) + CropBox Offset
                        // Rect.Bottom (Y1) = (PageHeight + OffsetY) - (WPF_Y + WPF_Height)
                        double pdfY = (pageHeight + yOffset) - (ann.Y + ann.Height);
                        double pdfX = xOffset + ann.X;

                        // 명시적으로 Point 지정하여 사각형 생성 (좌하단, 우상단)
                        var rect = new PdfRectangle(new XPoint(pdfX, pdfY), new XPoint(pdfX + ann.Width, pdfY + ann.Height));

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
                            
                            // [보정] 폰트 크기: WPF(96 DPI) -> PDF(72 DPI) = 0.75배
                            double pdfFontSize = ann.FontSize * 72.0 / 96.0;

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
                                // 여백 없이(0,0) 그리기
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
                            // 형광펜 / 밑줄
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

                            // [중요] QuadPoints 설정 (PDF 표준 순서: LT, RT, LB, RB 가 아니라 Z order)
                            // PDF QuadPoints는 8개의 숫자로 구성됨: x1, y1, x2, y2, x3, y3, x4, y4
                            // 순서: Top-Left, Top-Right, Bottom-Left, Bottom-Right 
                            // *주의: PDF 좌표계이므로 Y값이 클수록 위쪽임*
                            
                            double x = pdfX;
                            double y_bottom = pdfY;
                            double y_top = pdfY + ann.Height;
                            double w = ann.Width;

                            var qp = new PdfArray(doc);
                            qp.Elements.Add(new PdfReal(x));     qp.Elements.Add(new PdfReal(y_top));    // Top-Left
                            qp.Elements.Add(new PdfReal(x + w)); qp.Elements.Add(new PdfReal(y_top));    // Top-Right
                            qp.Elements.Add(new PdfReal(x));     qp.Elements.Add(new PdfReal(y_bottom)); // Bottom-Left
                            qp.Elements.Add(new PdfReal(x + w)); qp.Elements.Add(new PdfReal(y_bottom)); // Bottom-Right
                            
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