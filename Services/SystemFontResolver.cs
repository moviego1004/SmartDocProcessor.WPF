using System;
using System.IO;
using PdfSharp.Fonts;

namespace SmartDocProcessor.WPF.Services
{
    public class SystemFontResolver : IFontResolver
    {
        public string DefaultFontName => "Malgun Gothic";

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // 폰트 이름 매핑 (단순화: 무조건 해당 이름으로 시도)
            // 스타일(Bold/Italic)은 시뮬레이션으로 처리
            return new FontResolverInfo(familyName, isBold, isItalic);
        }

        public byte[]? GetFont(string faceName)
        {
            // 1. 요청한 폰트 파일 찾기
            byte[]? fontData = FindFontFile(faceName);
            if (fontData != null) return fontData;

            // 2. Fallback: 맑은 고딕
            fontData = FindFontFile("Malgun Gothic"); 
            if (fontData != null) return fontData;
            
            fontData = FindFontFile("malgun.ttf"); 
            if (fontData != null) return fontData;

            // 3. Fallback: Arial
            fontData = FindFontFile("Arial");
            if (fontData != null) return fontData;
            
            fontData = FindFontFile("arial.ttf");
            if (fontData != null) return fontData;

            // 4. 정말 아무것도 없다면? (거의 불가능하지만 크래시 방지용)
            // 에러를 던지지 말고 null을 리턴하면 PDFSharp이 내부 기본 폰트를 쓰거나 예외를 던짐
            // 여기서는 상위 PdfService의 try-catch가 처리하도록 null 반환
            return null;
        }

        private byte[]? FindFontFile(string name)
        {
            try
            {
                string fontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
                
                // 1. 정확한 이름 매칭 시도
                string fullPath = Path.Combine(fontsPath, name);
                if (File.Exists(fullPath)) return File.ReadAllBytes(fullPath);

                // 2. 확장자 붙여서 시도 (.ttf, .ttc)
                if (File.Exists(fullPath + ".ttf")) return File.ReadAllBytes(fullPath + ".ttf");
                if (File.Exists(fullPath + ".ttc")) return File.ReadAllBytes(fullPath + ".ttc");

                // 3. 특수 매핑
                if (name.Contains("Malgun", StringComparison.OrdinalIgnoreCase)) 
                    return ReadIfExists(Path.Combine(fontsPath, "malgun.ttf"));
                
                if (name.Contains("Gulim", StringComparison.OrdinalIgnoreCase) || name.Contains("Dotum", StringComparison.OrdinalIgnoreCase)) 
                    return ReadIfExists(Path.Combine(fontsPath, "gulim.ttc"));
                
                if (name.Contains("Batang", StringComparison.OrdinalIgnoreCase)) 
                    return ReadIfExists(Path.Combine(fontsPath, "batang.ttc"));
            }
            catch { }
            return null;
        }

        private byte[]? ReadIfExists(string path)
        {
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }
    }
}