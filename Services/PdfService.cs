using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions; // [필수] Regex
using System.Globalization; // [필수] CultureInfo, NumberStyles
using System.Reflection; // [필수] BindingFlags
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Fonts; 
using PdfSharp.Pdf.Annotations; 
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Drawing.Layout; // [필수] XTextFormatter
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace SmartDocProcessor.WPF.Services
{
    // [참고] TextData 클래스는 OcrService 등에 정의되어 있다고 가정

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

    // 상태 추적 클래스 (내부 정의)
    public class TextExtractionState
    {
        public TextState Current { get; set; } = new TextState();
        private Stack<TextState> _stack = new Stack<TextState>();

        public void Push() { _stack.Push(Current.Clone()); }
        public void Pop() { if (_stack.Count > 0) Current = _stack.Pop(); }
    }

    public class TextState
    {
        public XMatrix CTM { get; set; } = new XMatrix(1, 0, 0, 1, 0, 0);
        public XMatrix Tm { get; set; } = new XMatrix(1, 0, 0, 1, 0, 0);
        public XMatrix Tlm { get; set; } = new XMatrix(1, 0, 0, 1, 0, 0);
        public double FontSize { get; set; } = 10;

        public TextState Clone()
        {
            return new TextState 
            { 
                CTM = this.CTM, 
                Tm = this.Tm, 
                Tlm = this.Tlm, 
                FontSize = this.FontSize 
            };
        }
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
                if (op.OpCode.Name == "Do") return true; 
            } else if (content is CSequence seq) {
                foreach (var item in seq) if (HasTextOperators(item)) return true;
            }
            return false;
        }

        // [핵심] 텍스트 및 좌표 추출
        public List<TextData> ExtractTextFromPage(byte[] pdfBytes, int pageIndex)
        {
            var result = new List<TextData>();
            try
            {
                using (var inStream = new MemoryStream(pdfBytes))
                {
                    var doc = PdfReader.Open(inStream, PdfDocumentOpenMode.Import);
                    if (pageIndex < 1 || pageIndex > doc.PageCount) return result;

                    var page = doc.Pages[pageIndex - 1];
                    
                    // [수정] .Point 제거 (XUnit은 double로 자동 변환됨 or 이미 double임)
                    // PdfSharp 최신 버전에서는 XUnit.Point 프로퍼티가 Obsolete일 수 있으므로 
                    // 그냥 double로 받거나, 명시적 형변환을 사용. 여기선 안전하게 double로 간주.
                    double pageHeight = page.CropBox.Height > 0 ? (double)page.CropBox.Height : (double)page.Height;
                    double xOffset = (double)page.CropBox.X1;
                    double yOffset = (double)page.CropBox.Y1;

                    var state = new TextExtractionState();
                    // 초기값 설정
                    state.Current.CTM = new XMatrix(1, 0, 0, 1, 0, 0); 
                    state.Current.Tm = new XMatrix(1, 0, 0, 1, 0, 0);

                    var content = ContentReader.ReadContent(page);
                    ExtractTextRecursively(content, result, state, pageHeight, xOffset, yOffset, page.Resources);
                }
            }
            catch { }
            return result;
        }

        private void ExtractTextRecursively(CObject content, List<TextData> result, TextExtractionState state, double pageHeight, double xOffset, double yOffset, PdfResources resources)
        {
            if (content is CSequence seq)
            {
                foreach (var item in seq) ExtractTextRecursively(item, result, state, pageHeight, xOffset, yOffset, resources);
            }
            else if (content is COperator op)
            {
                if (op.OpCode.Name == "q") state.Push();
                else if (op.OpCode.Name == "Q") state.Pop();
                else if (op.OpCode.Name == "cm")
                {
                    if (op.Operands.Count >= 6)
                    {
                        var mat = new XMatrix(
                            GetOperandValue(op.Operands[0]), GetOperandValue(op.Operands[1]),
                            GetOperandValue(op.Operands[2]), GetOperandValue(op.Operands[3]),
                            GetOperandValue(op.Operands[4]), GetOperandValue(op.Operands[5])
                        );
                        state.Current.CTM.Prepend(mat);
                    }
                }
                else if (op.OpCode.Name == "Do")
                {
                    if (op.Operands.Count > 0 && op.Operands[0] is CName xObjName && resources != null && resources.XObjects != null)
                    {
                        var xObjectItem = resources.XObjects[xObjName.Value];
                        if (xObjectItem is PdfReference refObj) xObjectItem = refObj.Value;

                        if (xObjectItem is PdfDictionary xObjDict && xObjDict.Elements.GetString("/Subtype") == "/Form")
                        {
                            try 
                            {
                                if (xObjDict.Stream != null)
                                {
                                    byte[] formBytes = xObjDict.Stream.UnfilteredValue;
                                    var formContent = ContentReader.ReadContent(formBytes);
                                    
                                    var formResItem = xObjDict.Elements["/Resources"];
                                    PdfResources formResources = resources; 
                                    
                                    if (formResItem is PdfReference resRef) formResItem = resRef.Value;
                                    if (formResItem is PdfDictionary resDict) formResources = new PdfResources(resDict);

                                    state.Push();
                                    ExtractTextRecursively(formContent, result, state, pageHeight, xOffset, yOffset, formResources);
                                    state.Pop();
                                }
                            }
                            catch { }
                        }
                    }
                }
                else if (op.OpCode.Name == "BT")
                {
                    state.Current.Tm = new XMatrix(1, 0, 0, 1, 0, 0);
                    state.Current.Tlm = new XMatrix(1, 0, 0, 1, 0, 0);
                }
                else if (op.OpCode.Name == "Tf")
                {
                    if (op.Operands.Count >= 2) state.Current.FontSize = GetOperandValue(op.Operands[1]);
                }
                else if (op.OpCode.Name == "Tm")
                {
                    if (op.Operands.Count >= 6)
                    {
                        state.Current.Tm = new XMatrix(
                            GetOperandValue(op.Operands[0]), GetOperandValue(op.Operands[1]),
                            GetOperandValue(op.Operands[2]), GetOperandValue(op.Operands[3]),
                            GetOperandValue(op.Operands[4]), GetOperandValue(op.Operands[5])
                        );
                        state.Current.Tlm = state.Current.Tm;
                    }
                }
                else if (op.OpCode.Name == "Td" || op.OpCode.Name == "TD")
                {
                    if (op.Operands.Count >= 2)
                    {
                        double tx = GetOperandValue(op.Operands[0]);
                        double ty = GetOperandValue(op.Operands[1]);
                        state.Current.Tlm.TranslatePrepend(tx, ty);
                        state.Current.Tm = state.Current.Tlm;
                    }
                }
                else if (op.OpCode.Name == "T*")
                {
                    state.Current.Tlm.TranslatePrepend(0, -state.Current.FontSize * 1.2);
                    state.Current.Tm = state.Current.Tlm;
                }
                else if (op.OpCode.Name == "Tj" || op.OpCode.Name == "\'")
                {
                    if (op.Operands.Count > 0 && op.Operands[0] is CString cStr)
                        AddTextResult(result, cStr.Value, state.Current, pageHeight, xOffset, yOffset);
                }
                else if (op.OpCode.Name == "TJ")
                {
                    if (op.Operands.Count > 0 && op.Operands[0] is CArray arr)
                    {
                        var sb = new StringBuilder();
                        foreach (var item in arr) if (item is CString s) sb.Append(s.Value);
                        AddTextResult(result, sb.ToString(), state.Current, pageHeight, xOffset, yOffset);
                    }
                }
            }
        }

        private double GetOperandValue(CObject obj)
        {
            if (obj is CReal r) return r.Value;
            if (obj is CInteger i) return i.Value;
            return 0;
        }

        private void AddTextResult(List<TextData> result, string text, TextState current, double pageHeight, double xOffset, double yOffset)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            var mat = current.Tm * current.CTM;
            double pdfX = mat.OffsetX;
            double pdfY = mat.OffsetY;

            double wpfX = pdfX - xOffset;
            double wpfY = (pageHeight - (pdfY - yOffset)) - current.FontSize;

            double scaleY = Math.Sqrt(mat.M21 * mat.M21 + mat.M22 * mat.M22);
            double actualFontSize = current.FontSize * (scaleY == 0 ? 1 : scaleY);

            double widthPerChar = actualFontSize * 0.6; 
            double estimatedWidth = text.Length * widthPerChar;
            
            result.Add(new TextData
            {
                Text = text,
                X = wpfX / 0.75,
                Y = wpfY / 0.75,
                Width = estimatedWidth / 0.75,
                Height = actualFontSize / 0.75
            });
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
                        
                        // [수정] .Point 제거 및 명시적 캐스팅
                        double pageHeight = page.CropBox.Height > 0 ? (double)page.CropBox.Height : (double)page.Height;
                        double yOffset = (double)page.CropBox.Y1;
                        double xOffset = (double)page.CropBox.X1;

                        if (page.Annotations != null) {
                            foreach (var item in page.Annotations) {
                                if (item is PdfAnnotation annot) {
                                    var rect = annot.Rectangle;
                                    double appY = (pageHeight + yOffset) - rect.Y2;
                                    double appX = rect.X1 - xOffset;
                                    var data = new AnnotationData { X=appX/0.75, Y=appY/0.75, Width=rect.Width/0.75, Height=rect.Height/0.75, Page=pageNum, Content=annot.Contents??"", Type="TEXT" };
                                    var subType = annot.Elements.GetString("/Subtype");
                                    if(subType=="/Highlight"){ data.Type="HIGHLIGHT_Y"; data.Color="#FFFF00"; result.Add(data); }
                                    else if(subType=="/Underline"){ data.Type="UNDERLINE"; data.Color="#FF0000"; result.Add(data); }
                                    else if(subType=="/FreeText"){ ParseAppearanceString(annot.Elements.GetString("/DA"), data); result.Add(data); }
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
                if (fontMatch.Success && double.TryParse(fontMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double size)) data.FontSize = (int)Math.Round(size / 0.75);
                var rgbMatch = Regex.Match(da, @"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+rg");
                if (rgbMatch.Success) {
                    double r=double.Parse(rgbMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    double g=double.Parse(rgbMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                    double b=double.Parse(rgbMatch.Groups[3].Value, CultureInfo.InvariantCulture);
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
                    
                    // [수정] .Point 제거 및 명시적 캐스팅
                    double pageHeight = page.CropBox.Height > 0 ? (double)page.CropBox.Height : (double)page.Height;
                    double pageTop = page.CropBox.Height > 0 ? ((double)page.CropBox.Y1 + (double)page.CropBox.Height) : (double)page.Height;
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
                        double scaledX = ann.X * 0.75; double scaledY = ann.Y * 0.75;
                        double scaledW = Math.Max(ann.Width * 0.75, 1.0); double scaledH = Math.Max(ann.Height * 0.75, 1.0);
                        double pdfTopY = pageTop - scaledY; double pdfBottomY = pdfTopY - scaledH; double pdfLeftX = xOffset + scaledX;
                        var rect = new PdfRectangle(new XPoint(pdfLeftX, pdfBottomY), new XPoint(pdfLeftX + scaledW, pdfTopY));

                        if (ann.Type == "TEXT") {
                            var annot = new CustomPdfAnnotation(doc);
                            annot.Elements.SetName(PdfAnnotation.Keys.Type, "/Annot");
                            annot.Elements.SetName(PdfAnnotation.Keys.Subtype, "/FreeText");
                            annot.Elements.SetRectangle(PdfAnnotation.Keys.Rect, rect);
                            annot.Elements.SetString(PdfAnnotation.Keys.Contents, ann.Content);
                            XColor c = XColor.FromArgb(255,0,0,0);
                            try { string hex=ann.Color.Replace("#",""); if(hex.Length==8) c=XColor.FromArgb(int.Parse(hex.Substring(0,2),NumberStyles.HexNumber), int.Parse(hex.Substring(2,2),NumberStyles.HexNumber), int.Parse(hex.Substring(4,2),NumberStyles.HexNumber), int.Parse(hex.Substring(6,2),NumberStyles.HexNumber)); } catch{}
                            double r=c.R/255.0; double g=c.G/255.0; double b=c.B/255.0; double pdfFontSize=ann.FontSize*0.75;
                            annot.Elements["/DA"] = new PdfString($"/Helv {pdfFontSize:0.##} Tf {r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b.ToString("0.###", CultureInfo.InvariantCulture)} rg");
                            var formRect = new XRect(0, 0, scaledW, scaledH);
                            var form = new XForm(doc, formRect);
                            using (var gfx = XGraphics.FromForm(form)) {
                                var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                var font = CreateSafeFont(ann.FontFamily, Math.Max(pdfFontSize,4), ann.IsBold?XFontStyleEx.Bold:XFontStyleEx.Regular, options);
                                var brush = new XSolidBrush(c);
                                var tf = new XTextFormatter(gfx); tf.Alignment=XParagraphAlignment.Left;
                                tf.DrawString(ann.Content, font, brush, formRect, XStringFormats.TopLeft);
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

        private XFont CreateSafeFont(string familyName, double size, XFontStyleEx style, XPdfFontOptions? options = null) {
            if (options == null) options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            try { return new XFont(familyName, size, style, options); }
            catch { try { return new XFont("Malgun Gothic", size, style, options); } catch { return new XFont("Arial", size, style, options); } }
        }
        public byte[] DeletePage(byte[] pdfBytes, int pageIndex) { using (var ms = new MemoryStream(pdfBytes)) using (var outMs = new MemoryStream()) { var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Import); var newDoc = new PdfDocument(); for (int i = 0; i < doc.PageCount; i++) if (i != pageIndex) newDoc.AddPage(doc.Pages[i]); newDoc.Save(outMs); return outMs.ToArray(); } }
        private PdfDictionary? GetPdfForm(XForm form) { try { var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public); if (prop != null) return prop.GetValue(form) as PdfDictionary; } catch { } return null; }
    }

    public class CustomPdfAnnotation : PdfAnnotation { public CustomPdfAnnotation(PdfDocument document) : base(document) { } }
}