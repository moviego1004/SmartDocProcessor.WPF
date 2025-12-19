using System.Collections.Generic;
using System.IO;

namespace SmartDocProcessor.WPF.Services
{
    public class PdfDocumentModel
    {
        public string FilePath { get; set; } = "";
        public string FileName => System.IO.Path.GetFileName(FilePath);
        
        // [신규] 텍스트 추출 전용 원본 데이터 (저장 후에도 변하지 않음)
        public byte[]? OriginalPdfData { get; set; }

        // 편집/저장용 데이터
        public byte[]? PdfData { get; set; }
        public byte[]? CleanPdfData { get; set; }
        
        // 주석 및 OCR 데이터
        public List<AnnotationData> Annotations { get; set; } = new List<AnnotationData>();
        public Dictionary<int, List<TextData>> PageTextData { get; set; } = new Dictionary<int, List<TextData>>();
        
        // 뷰어 상태
        public int TotalPages { get; set; } = 0;
        public int CurrentPage { get; set; } = 1;
        public double ZoomLevel { get; set; } = 1.0;
        
        // 스크롤 위치
        public double VerticalOffset { get; set; } = 0;
        public double HorizontalOffset { get; set; } = 0;
        
        public bool IsSelected { get; set; }
    }
}