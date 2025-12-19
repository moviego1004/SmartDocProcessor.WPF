using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Fonts;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.Advanced;

// PdfPig 별칭
using PigPdfDocument = UglyToad.PdfPig.PdfDocument;

namespace SmartDocProcessor.WPF.Services
{
    // TextData 클래스는 OcrService에 정의됨

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
        public PdfService()
        {
            try { if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new SystemFontResolver(); } catch { }
        }

        // [읽기] PdfPig를 사용하여 Searchable 여부 확인
        public bool IsPdfSearchable(byte[] pdfBytes)
        {
            try
            {
                using (var doc = PigPdfDocument.Open(pdfBytes))
                {
                    if (doc.NumberOfPages == 0) return false;
                    int pagesToCheck = Math.Min(doc.NumberOfPages, 3);
                    for (int i = 1; i <= pagesToCheck; i++)
                    {
                        var page = doc.GetPage(i);
                        if (page.GetWords().Any()) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        // [핵심 기능] 텍스트 및 정확한 좌표 추출 (스케일링 없이 순수 좌표 반환)
        public List<TextData> ExtractTextFromPage(byte[] pdfBytes, int pageIndex)
        {
            var result = new List<TextData>();
            try
            {
                using (var doc = PigPdfDocument.Open(pdfBytes))
                {
                    if (pageIndex < 1 || pageIndex > doc.NumberOfPages) return result;

                    var page = doc.GetPage(pageIndex);
                    
                    var cropBox = page.CropBox.Bounds;
                    int rotation = page.Rotation.Value;
                    rotation = (rotation % 360 + 360) % 360; // 정규화

                    foreach (var word in page.GetWords())
                    {
                        if (string.IsNullOrWhiteSpace(word.Text)) continue;

                        var box = word.BoundingBox;
                        
                        // 1. CropBox 기준 상대 좌표 계산 (PDF 좌표계: Bottom-Left 기준)
                        // 단어의 절대 좌표에서 페이지 시작점(여백)을 뺍니다.
                        double relLeft = box.Left - cropBox.Left;
                        double relBottom = box.Bottom - cropBox.Bottom;
                        
                        double wpfX = 0, wpfY = 0, wpfW = 0, wpfH = 0;

                        // 회전된 페이지의 크기 (Points)
                        double pageWidth = (rotation == 90 || rotation == 270) ? cropBox.Height : cropBox.Width;
                        double pageHeight = (rotation == 90 || rotation == 270) ? cropBox.Width : cropBox.Height;

                        // 2. 회전 변환 (Top-Left 기준 시각적 좌표로 변환)
                        switch (rotation)
                        {
                            case 0:
                                // 0도: Y축 반전 (Bottom-Up -> Top-Down)
                                wpfX = relLeft;
                                wpfY = pageHeight - (relBottom + box.Height);
                                wpfW = box.Width;
                                wpfH = box.Height;
                                break;

                            case 90: // 시계방향 90도
                                // Bottom -> Visual X, Left -> Visual Y
                                wpfX = relBottom; 
                                wpfY = relLeft; 
                                wpfW = box.Height;
                                wpfH = box.Width;
                                break;

                            case 180: // 상하좌우 반전
                                wpfX = pageWidth - (relLeft + box.Width);
                                wpfY = relBottom; 
                                wpfW = box.Width;
                                wpfH = box.Height;
                                break;

                            case 270: // 반시계 90도
                                wpfX = pageWidth - (relBottom + box.Height);
                                wpfY = pageHeight - (relLeft + box.Width);
                                wpfW = box.Height;
                                wpfH = box.Width;
                                break;
                        }

                        // [수정] 배율 곱하기 제거! (순수 PDF Point 좌표 반환)
                        // MainWindow에서 렌더링 시 1.5배를 곱하므로 여기서 또 곱하면 안 됨.

                        result.Add(new TextData
                        {
                            Text = word.Text,
                            X = wpfX,
                            Y = wpfY,
                            Width = wpfW,
                            Height = wpfH
                        });
                    }
                }
            }
            catch { }
            return result;
        }

        // --- 이하 저장/주석 로드 (기존 로직 유지) ---

        public List<AnnotationData> ExtractAnnotationsFromMetadata(byte[] pdfBytes)
        {
            var result = new List<AnnotationData>();
            try
            {
                using (var inStream = new MemoryStream(pdfBytes))
                {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Import);
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        int pageNum = i + 1;
                        
                        double pageHeight = (double)page.Height;
                        if ((double)page.CropBox.Height > 0) pageHeight = (double)page.CropBox.Height;
                        
                        double yOffset = (double)page.CropBox.Y1;
                        double xOffset = (double)page.CropBox.X1;

                        if (page.Annotations != null)
                        {
                            foreach (var item in page.Annotations)
                            {
                                if (item is PdfAnnotation annot)
                                {
                                    var rect = annot.Rectangle;
                                    
                                    // 기존 PdfSharp 좌표계(72dpi) -> 화면 좌표계(96dpi)로 변환 (기존 유지)
                                    // 여기는 1.33배 해줘야 뷰어에서 1.5배 할 때 총 2.0배가 되어 적절함 (기존 로직 따름)
                                    double appY = (pageHeight + yOffset) - rect.Y2;
                                    double appX = rect.X1 - xOffset;
                                    
                                    var data = new AnnotationData 
                                    { 
                                        X = appX / 0.75, 
                                        Y = appY / 0.75, 
                                        Width = rect.Width / 0.75, 
                                        Height = rect.Height / 0.75, 
                                        Page = pageNum, 
                                        Content = annot.Contents ?? "", 
                                        Type = "TEXT" 
                                    };

                                    var subType = annot.Elements.GetString("/Subtype");
                                    if (subType == "/Highlight") { data.Type = "HIGHLIGHT_Y"; data.Color = "#FFFF00"; result.Add(data); }
                                    else if (subType == "/Underline") { data.Type = "UNDERLINE"; data.Color = "#FF0000"; result.Add(data); }
                                    else if (subType == "/FreeText") 
                                    { 
                                        data.Type = "TEXT"; 
                                        ParseAppearanceString(annot.Elements.GetString("/DA"), data); 
                                        result.Add(data); 
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return result;
        }

        private void ParseAppearanceString(string da, AnnotationData data)
        {
            if (string.IsNullOrEmpty(da)) return;
            try
            {
                var fontMatch = Regex.Match(da, @"([\d\.]+)\s+Tf");
                double size = 0;
                if (fontMatch.Success && double.TryParse(fontMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out size)) 
                {
                    data.FontSize = (int)Math.Round(size / 0.75);
                }

                var rgbMatch = Regex.Match(da, @"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+rg");
                if (rgbMatch.Success)
                {
                    double r = double.Parse(rgbMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    double g = double.Parse(rgbMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                    double b = double.Parse(rgbMatch.Groups[3].Value, CultureInfo.InvariantCulture);
                    data.Color = $"#FF{(int)(r * 255):X2}{(int)(g * 255):X2}{(int)(b * 255):X2}";
                }
            }
            catch { }
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
                    
                    double pageHeight = (double)page.Height;
                    if ((double)page.CropBox.Height > 0) pageHeight = (double)page.CropBox.Height;
                    double pageTop = (double)page.Height;
                    if ((double)page.CropBox.Height > 0) pageTop = (double)page.CropBox.Y1 + (double)page.CropBox.Height;
                    double xOffset = (double)page.CropBox.X1;

                    var ocrItems = group.Where(a => a.Type == "OCR_TEXT");
                    if (ocrItems.Any())
                    {
                        using (var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append))
                        {
                            foreach (var ann in ocrItems)
                            {
                                var brush = new XSolidBrush(XColor.FromArgb(1, 0, 0, 0));
                                var font = new XFont("Arial", 10, XFontStyleEx.Regular);
                                gfx.DrawString(ann.Content, font, brush, new XRect(ann.X, ann.Y, ann.Width, ann.Height), XStringFormats.TopLeft);
                            }
                        }
                    }

                    foreach (var ann in group.Where(a => a.Type != "OCR_TEXT"))
                    {
                        double scaledX = ann.X * 0.75; 
                        double scaledY = ann.Y * 0.75;
                        double scaledW = Math.Max(ann.Width * 0.75, 1.0); 
                        double scaledH = Math.Max(ann.Height * 0.75, 1.0);

                        double pdfTopY = pageTop - scaledY;
                        double pdfBottomY = pdfTopY - scaledH;
                        double pdfLeftX = xOffset + scaledX;

                        var rect = new PdfRectangle(new XPoint(pdfLeftX, pdfBottomY), new XPoint(pdfLeftX + scaledW, pdfTopY));

                        if (ann.Type == "TEXT") {
                            var annot = new CustomPdfAnnotation(doc);
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, "/FreeText");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            annot.Elements.SetString(PdfAnnotation.Keys.Contents, ann.Content);
                            XColor c = XColor.FromArgb(255,0,0,0);
                            try { string hex=ann.Color.Replace("#",""); if(hex.Length==8) c=XColor.FromArgb(int.Parse(hex.Substring(0,2),NumberStyles.HexNumber), int.Parse(hex.Substring(2,2),NumberStyles.HexNumber), int.Parse(hex.Substring(4,2),NumberStyles.HexNumber), int.Parse(hex.Substring(6,2),NumberStyles.HexNumber)); } catch{}
                            double r=c.R/255.0; double g=c.G/255.0; double b=c.B/255.0; 
                            double pdfFontSize = ann.FontSize * 0.75;

                            annot.Elements["/DA"] = new PdfString($"/Helv {pdfFontSize:0.##} Tf {r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b.ToString("0.###", CultureInfo.InvariantCulture)} rg");
                            var formRect = new XRect(0, 0, scaledW, scaledH);
                            var form = new XForm(doc, formRect);
                            using (var gfx = XGraphics.FromForm(form)) {
                                var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                var font = CreateSafeFont(ann.FontFamily, Math.Max(pdfFontSize,4), ann.IsBold?XFontStyleEx.Bold:XFontStyleEx.Regular, options);
                                var brush = new XSolidBrush(c);
                                gfx.DrawString(ann.Content, font, brush, formRect, XStringFormats.TopLeft);
                            }
                            var apDict = new PdfDictionary(doc); annot.Elements["/AP"]=apDict;
                            var pdfForm = GetPdfForm(form); if(pdfForm!=null) apDict.Elements["/N"]=pdfForm;
                            page.Annotations.Add(annot);
                        } else {
                            var annot = new CustomPdfAnnotation(doc);
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, ann.Type.StartsWith("HIGHLIGHT")?"/Highlight":"/Underline");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            if(ann.Type.StartsWith("HIGHLIGHT")) {
                                var color=ann.Type=="HIGHLIGHT_O"?XColor.FromArgb(255,165,0):XColor.FromArgb(255,255,0);
                                annot.Elements["/C"]=new PdfArray(doc, new PdfReal(color.R/255.0), new PdfReal(color.G/255.0), new PdfReal(color.B/255.0));
                            } else { annot.Elements["/C"]=new PdfArray(doc, new PdfReal(1), new PdfReal(0), new PdfReal(0)); }
                            var qp = new PdfArray(doc);
                            qp.Elements.Add(new PdfReal(pdfLeftX)); qp.Elements.Add(new PdfReal(pdfTopY));
                            qp.Elements.Add(new PdfReal(pdfLeftX+scaledW)); qp.Elements.Add(new PdfReal(pdfTopY));
                            qp.Elements.Add(new PdfReal(pdfLeftX)); qp.Elements.Add(new PdfReal(pdfBottomY));
                            qp.Elements.Add(new PdfReal(pdfLeftX+scaledW)); qp.Elements.Add(new PdfReal(pdfBottomY));
                            annot.Elements["/QuadPoints"] = qp;
                            page.Annotations.Add(annot);
                        }
                    }
                }
                doc.Save(outStream);
                return outStream.ToArray();
            }
        }

        public byte[] DeletePage(byte[] pdfBytes, int pageIndex) { using (var ms = new MemoryStream(pdfBytes)) using (var outMs = new MemoryStream()) { var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import); var newDoc = new PdfDocument(); for (int i = 0; i < doc.PageCount; i++) if (i != pageIndex) newDoc.AddPage(doc.Pages[i]); newDoc.Save(outMs); return outMs.ToArray(); } }
        
        private PdfDictionary? GetPdfForm(XForm form) { try { var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public); if (prop != null) return prop.GetValue(form) as PdfDictionary; } catch { } return null; }
        
        private XFont CreateSafeFont(string familyName, double size, XFontStyleEx style, XPdfFontOptions? options = null) {
            if (options == null) options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            try { return new XFont(familyName, size, style, options); }
            catch { try { return new XFont("Malgun Gothic", size, style, options); } catch { return new XFont("Arial", size, style, options); } }
        }
    }

    public class CustomPdfAnnotation : PdfAnnotation { public CustomPdfAnnotation(PdfDocument document) : base(document) { } }
}