using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Reflection; // BindingFlags
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

        // [신규] PDF가 Searchable(텍스트 포함)인지 확인
        public bool IsPdfSearchable(byte[] pdfBytes)
        {
            try
            {
                using (var inStream = new MemoryStream(pdfBytes))
                {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Import);
                    int pagesToCheck = Math.Min(doc.PageCount, 3);
                    for (int i = 0; i < pagesToCheck; i++)
                    {
                        var content = ContentReader.ReadContent(doc.Pages[i]);
                        if (HasTextOperators(content)) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private bool HasTextOperators(CObject content)
        {
            if (content is COperator op)
            {
                if (op.OpCode.Name == "Tj" || op.OpCode.Name == "TJ" || op.OpCode.Name == "\'") return true;
            }
            else if (content is CSequence seq)
            {
                foreach (var item in seq) if (HasTextOperators(item)) return true;
            }
            return false;
        }

        public byte[] GetPdfBytesWithoutAnnotations(byte[] pdfBytes)
        {
            try
            {
                using (var inStream = new MemoryStream(pdfBytes))
                using (var outStream = new MemoryStream())
                {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Modify);
                    foreach (var page in doc.Pages) if (page.Annotations != null) page.Annotations.Clear();
                    doc.Save(outStream);
                    return outStream.ToArray();
                }
            }
            catch { return pdfBytes; }
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
                        double pageHeight = page.Height.Point;
                        if (page.Annotations != null) {
                            foreach (var item in page.Annotations) {
                                if (item is PdfAnnotation annot) {
                                    var subType = annot.Elements.GetString("/Subtype");
                                    var rect = annot.Rectangle;
                                    double appY = pageHeight - rect.Y2; 
                                    var data = new AnnotationData { X=rect.X1, Y=appY, Width=rect.Width, Height=rect.Height, Page=pageNum, Content=annot.Contents ?? "" };
                                    
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
                if (fontMatch.Success && double.TryParse(fontMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double size)) data.FontSize = (int)Math.Round(size);
                
                var rgbMatch = Regex.Match(da, @"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+rg");
                if (rgbMatch.Success) {
                    double r = double.Parse(rgbMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    double g = double.Parse(rgbMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                    double b = double.Parse(rgbMatch.Groups[3].Value, CultureInfo.InvariantCulture);
                    data.Color = $"#FF{(int)(r*255):X2}{(int)(g*255):X2}{(int)(b*255):X2}";
                }
                else {
                    var grayMatch = Regex.Match(da, @"([\d\.]+)\s+g");
                    if (grayMatch.Success) {
                        double v = double.Parse(grayMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                        int gray = (int)(v * 255);
                        data.Color = $"#FF{gray:X2}{gray:X2}{gray:X2}";
                    }
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
                    double pageHeight = page.Height.Point;

                    foreach (var ann in group)
                    {
                        if (ann.Type == "OCR_TEXT") continue;

                        double pdfY = pageHeight - ann.Y - ann.Height;
                        var rect = new PdfRectangle(new XRect(ann.X, pdfY, ann.Width, ann.Height));

                        if (ann.Type == "TEXT")
                        {
                            var annot = new CustomPdfAnnotation(doc); // [수정] CustomPdfAnnotation 사용
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
                            annot.Elements["/DA"] = new PdfString($"/Helv {ann.FontSize} Tf {r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b.ToString("0.###", CultureInfo.InvariantCulture)} rg");

                            double formW = Math.Max(ann.Width, 10);
                            double formH = Math.Max(ann.Height, 10) + 10; 
                            var formRect = new XRect(0, 0, formW, formH);
                            var form = new XForm(doc, formRect);
                            
                            using (var gfx = XGraphics.FromForm(form))
                            {
                                var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                var font = CreateSafeFont(ann.FontFamily, Math.Max(ann.FontSize, 5), ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular, options);
                                var brush = new XSolidBrush(c);
                                var tf = new XTextFormatter(gfx);
                                tf.Alignment = XParagraphAlignment.Left;
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
                            var annot = new CustomPdfAnnotation(doc); // [수정] CustomPdfAnnotation 사용
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, ann.Type.StartsWith("HIGHLIGHT") ? "/Highlight" : "/Underline");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            
                            if(ann.Type.StartsWith("HIGHLIGHT")) {
                                var color = ann.Type == "HIGHLIGHT_O" ? XColor.FromArgb(255, 165, 0) : XColor.FromArgb(255, 255, 0);
                                annot.Elements["/C"] = new PdfArray(doc, new PdfReal(color.R / 255.0), new PdfReal(color.G / 255.0), new PdfReal(color.B / 255.0));
                            } else {
                                annot.Elements["/C"] = new PdfArray(doc, new PdfReal(1), new PdfReal(0), new PdfReal(0));
                            }
                            var qp = new PdfArray(doc);
                            qp.Elements.Add(new PdfReal(rect.X1)); qp.Elements.Add(new PdfReal(rect.Y2)); 
                            qp.Elements.Add(new PdfReal(rect.X2)); qp.Elements.Add(new PdfReal(rect.Y2)); 
                            qp.Elements.Add(new PdfReal(rect.X1)); qp.Elements.Add(new PdfReal(rect.Y1)); 
                            qp.Elements.Add(new PdfReal(rect.X2)); qp.Elements.Add(new PdfReal(rect.Y1)); 
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

    // [중요] CS0122 에러 해결을 위해 public 클래스로 선언
    public class CustomPdfAnnotation : PdfAnnotation
    {
        public CustomPdfAnnotation(PdfDocument document) : base(document) { }
    }
}