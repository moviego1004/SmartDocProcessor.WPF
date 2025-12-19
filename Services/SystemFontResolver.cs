using System;
using System.IO;
using PdfSharp.Fonts;

namespace SmartDocProcessor.WPF.Services
{
    // PDFSharp 6.x 이상에서 한글 폰트 사용을 위한 리졸버
    public class SystemFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
        {
            // 윈도우 폰트 폴더에서 폰트 파일 읽기 (간소화된 구현)
            // 실제 프로덕션에서는 더 정교한 매칭이 필요할 수 있음
            string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            string file = Path.Combine(fontPath, "malgun.ttf"); // 기본값: 맑은 고딕

            if (faceName.Contains("Arial", StringComparison.OrdinalIgnoreCase)) file = Path.Combine(fontPath, "arial.ttf");
            else if (faceName.Contains("Times", StringComparison.OrdinalIgnoreCase)) file = Path.Combine(fontPath, "times.ttf");
            // 다른 폰트 매핑 추가 가능

            if (File.Exists(file)) return File.ReadAllBytes(file);
            
            // 파일이 없으면 맑은 고딕 시도
            file = Path.Combine(fontPath, "malgun.ttf");
            if (File.Exists(file)) return File.ReadAllBytes(file);

            return new byte[0]; // 실패 시
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // 어떤 폰트 요청이 와도 기본적으로 처리
            return new FontResolverInfo(familyName);
        }
    }
}