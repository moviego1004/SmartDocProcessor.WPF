using System.Collections.Generic;
using System.IO;

namespace SmartDocProcessor.WPF.Services
{
    public class PdfDocumentModel
    {
        public string FilePath { get; set; } = "";
        public string FileName => System.IO.Path.GetFileName(FilePath);
        
        // PDF 원본 및 렌더링용 데이터
        public byte[]? PdfData { get; set; }
        public byte[]? CleanPdfData { get; set; }
        
        // 주석 및 OCR 데이터
        public List<AnnotationData> Annotations { get; set; } = new List<AnnotationData>();
        public Dictionary<int, List<TextData>> PageTextData { get; set; } = new Dictionary<int, List<TextData>>();
        
        // 뷰어 상태
        public int TotalPages { get; set; } = 0;
        public int CurrentPage { get; set; } = 1;
        public double ZoomLevel { get; set; } = 1.0;
        
        // UI 탭 선택용 (선택적)
        public bool IsSelected { get; set; }
    }
}