using System.Windows;
using PdfSharp.Fonts;
using SmartDocProcessor.WPF.Services;

namespace SmartDocProcessor.WPF
{
    public partial class App : Application
    {
        public App()
        {
            // [핵심] 한글 폰트 지원을 위해 FontResolver 등록
            if (GlobalFontSettings.FontResolver == null)
            {
                GlobalFontSettings.FontResolver = new SystemFontResolver();
            }
        }
    }
}