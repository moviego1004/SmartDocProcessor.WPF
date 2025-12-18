using Microsoft.Win32;
using SmartDocProcessor.WPF.Services; 
using SmartDocProcessor.WPF.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;
using System.Threading.Tasks;

namespace SmartDocProcessor.WPF
{
    public partial class MainWindow : Window
    {
        private readonly PdfService _pdfService = new PdfService();
        private readonly OcrService _ocrService = new OcrService();
        private UserSettings _userSettings;

        private string? _currentFilePath;
        private byte[]? _pdfData; 
        private byte[]? _cleanPdfData; 
        private List<AnnotationData> _annotations = new List<AnnotationData>();
        
        private Dictionary<int, List<TextData>> _pageTextData = new Dictionary<int, List<TextData>>();
        private List<TextData> _selectedTextData = new List<TextData>();

        private int _focusedPage = 1;
        private int _totalPages = 0;
        private string _currentTool = "CURSOR";
        private AnnotationData? _selectedAnnotation;
        
        private bool _isDrawing = false;
        private bool _isDraggingAnnot = false;
        private bool _isResizingAnnot = false;
        private bool _isTextSelecting = false;
        
        private Point _dragStartPoint;
        private Point _annotStartPos;
        private Size _annotStartSize;
        private Canvas? _currentDrawingCanvas; 

        private const double RENDER_SCALE = 1.5; 
        private double _zoomLevel = 1.0;
        private bool _isUpdatingUi = false;

        public MainWindow()
        {
            InitializeComponent();
            _userSettings = UserSettings.Load();
            if (string.IsNullOrEmpty(_userSettings.DefaultFontFamily)) _userSettings.DefaultFontFamily = "Malgun Gothic";
        }

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" };
            if (dlg.ShowDialog() == true)
            {
                _currentFilePath = dlg.FileName;
                _pdfData = await File.ReadAllBytesAsync(_currentFilePath);
                
                _annotations = _pdfService.ExtractAnnotationsFromMetadata(_pdfData);
                _cleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(_pdfData);
                _pageTextData.Clear();

                using (var doc = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(_pdfData), PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)) {
                    _totalPages = doc.PageCount;
                }
                TxtPageInfo.Text = $"총 {_totalPages} 페이지";
                _zoomLevel = 1.0; ApplyZoom();
                await RenderDocument();

                if (_pdfService.IsPdfSearchable(_pdfData))
                {
                    await LoadTextDataForPage(_focusedPage);
                }
            }
        }

        private async void BtnOcr_Click(object sender, RoutedEventArgs e)
        {
            if (_pdfData == null) return;
            
            MessageBox.Show("전체 페이지 OCR 분석을 시작합니다. 완료되면 드래그가 가능해집니다.");
            this.Cursor = Cursors.Wait;

            for (int i = 1; i <= _totalPages; i++)
            {
                if (!_pageTextData.ContainsKey(i))
                {
                    var texts = await _ocrService.ExtractTextData(_pdfData, i);
                    if (texts.Count > 0) _pageTextData[i] = texts;
                }
            }

            this.Cursor = Cursors.Arrow;
            MessageBox.Show("OCR 완료! 이제 텍스트 드래그가 가능합니다.");
        }

        private async Task LoadTextDataForPage(int page)
        {
            if (!_pageTextData.ContainsKey(page) && _pdfData != null)
            {
                var texts = await _ocrService.ExtractTextData(_pdfData, page);
                if (texts.Count > 0) _pageTextData[page] = texts;
            }
        }

        private async System.Threading.Tasks.Task RenderDocument()
        {
            var dataToRender = _cleanPdfData ?? _pdfData;
            if (dataToRender == null) return;

            DocumentContainer.Children.Clear();
            for (int i = 1; i <= _totalPages; i++)
            {
                var bitmap = await PdfRenderer.RenderPageToBitmapAsync(dataToRender, i);
                if (bitmap == null) continue;

                var pageGrid = new Grid { Margin = new Thickness(0, 0, 0, 20) };
                pageGrid.Width = bitmap.PixelWidth;
                pageGrid.Height = bitmap.PixelHeight;

                var img = new Image { Source = bitmap, Stretch = Stretch.None };
                pageGrid.Children.Add(img);

                var canvas = new Canvas 
                { 
                    Background = Brushes.Transparent, 
                    Width = bitmap.PixelWidth, Height = bitmap.PixelHeight, 
                    Tag = i 
                };
                canvas.MouseDown += Canvas_MouseDown;
                canvas.MouseMove += Canvas_MouseMove;
                canvas.MouseUp += Canvas_MouseUp;

                pageGrid.Children.Add(canvas);
                DrawAnnotationsForPage(i, canvas);
                DocumentContainer.Children.Add(pageGrid);
            }
        }

        private void RefreshPageCanvas(int pageIndex)
        {
            if (pageIndex < 1 || pageIndex > DocumentContainer.Children.Count) return;
            if (DocumentContainer.Children[pageIndex - 1] is Grid pageGrid)
            {
                var canvas = pageGrid.Children.OfType<Canvas>().FirstOrDefault();
                if (canvas != null) DrawAnnotationsForPage(pageIndex, canvas);
            }
        }

        // 캔버스 마우스 이벤트
        private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Canvas canvas) return;
            
            if (canvas.Tag is int pageNum) 
            {
                _focusedPage = pageNum;
                if (_pdfService.IsPdfSearchable(_pdfData ?? Array.Empty<byte>()) && !_pageTextData.ContainsKey(pageNum))
                {
                     _ = LoadTextDataForPage(pageNum); 
                }
            }

            _dragStartPoint = e.GetPosition(canvas);
            _currentDrawingCanvas = canvas;
            canvas.CaptureMouse();

            if (_currentTool == "CURSOR" && _pageTextData.ContainsKey(_focusedPage))
            {
                _isTextSelecting = true;
                _selectedTextData.Clear();
                ClearSelectionVisuals(canvas);
                return;
            }

            if (_currentTool != "CURSOR") _isDrawing = true;
        }

        private void Canvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (_currentDrawingCanvas == null) return;
            var currentPoint = e.GetPosition(_currentDrawingCanvas);

            if (_isTextSelecting)
            {
                UpdateTextSelection(_currentDrawingCanvas, _dragStartPoint, currentPoint);
                return;
            }

            if (_isDraggingAnnot && _selectedAnnotation != null)
            {
                double deltaX = (currentPoint.X - _dragStartPoint.X) / RENDER_SCALE;
                double deltaY = (currentPoint.Y - _dragStartPoint.Y) / RENDER_SCALE;
                _selectedAnnotation.X = _annotStartPos.X + deltaX;
                _selectedAnnotation.Y = _annotStartPos.Y + deltaY;
                RefreshPageCanvas(_selectedAnnotation.Page);
            }
            else if (_isResizingAnnot && _selectedAnnotation != null)
            {
                double deltaX = (currentPoint.X - _dragStartPoint.X) / RENDER_SCALE;
                double deltaY = (currentPoint.Y - _dragStartPoint.Y) / RENDER_SCALE;
                double newW = Math.Max(10, _annotStartSize.Width + deltaX);
                double newH = Math.Max(10, _annotStartSize.Height + deltaY);
                _selectedAnnotation.Width = newW;
                _selectedAnnotation.Height = newH;
                RefreshPageCanvas(_selectedAnnotation.Page);
            }
        }

        private void Canvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            _currentDrawingCanvas?.ReleaseMouseCapture();

            if (_isTextSelecting)
            {
                _isTextSelecting = false;
                if (_selectedTextData.Count > 0) PopupTextMenu.IsOpen = true;
                _currentDrawingCanvas = null;
                return;
            }

            if (_isDraggingAnnot || _isResizingAnnot)
            {
                _isDraggingAnnot = false; _isResizingAnnot = false;
                _currentDrawingCanvas = null;
                UpdatePropertyPanel(); return;
            }

            if (!_isDrawing || _currentDrawingCanvas == null) return;
            _isDrawing = false;

            var canvas = _currentDrawingCanvas;
            var endPoint = e.GetPosition(canvas);
            _currentDrawingCanvas = null;

            double x = Math.Min(_dragStartPoint.X, endPoint.X);
            double y = Math.Min(_dragStartPoint.Y, endPoint.Y);
            double w = Math.Abs(endPoint.X - _dragStartPoint.X);
            double h = Math.Abs(endPoint.Y - _dragStartPoint.Y);

            if (w < 5 && h < 5 && _currentTool != "TEXT") return;

            int pageNum = (canvas.Tag as int?) ?? 1;

            var newAnn = new AnnotationData
            {
                Type = _currentTool,
                X = x / RENDER_SCALE,
                Y = y / RENDER_SCALE,
                Width = Math.Max(w / RENDER_SCALE, 100), 
                Height = Math.Max(h / RENDER_SCALE, 30),
                Page = pageNum,
            };

            if (_currentTool == "TEXT")
            {
                newAnn.Content = "텍스트 입력";
                newAnn.FontFamily = _userSettings.DefaultFontFamily;
                newAnn.FontSize = _userSettings.DefaultFontSize;
                newAnn.IsBold = _userSettings.DefaultIsBold;
                newAnn.Color = _userSettings.DefaultColor;
            }

            _annotations.Add(newAnn);
            if (_currentTool == "TEXT") SelectAnnotation(newAnn);
            else RefreshPageCanvas(pageNum);
        }

        private void UpdateTextSelection(Canvas canvas, Point start, Point end)
        {
            double x = Math.Min(start.X, end.X);
            double y = Math.Min(start.Y, end.Y);
            double w = Math.Abs(end.X - start.X);
            double h = Math.Abs(end.Y - start.Y);
            var dragRect = new Rect(x, y, w, h);

            ClearSelectionVisuals(canvas);
            _selectedTextData.Clear();

            if (_pageTextData.ContainsKey(_focusedPage))
            {
                foreach (var word in _pageTextData[_focusedPage])
                {
                    Rect wordRect = new Rect(word.X * RENDER_SCALE, word.Y * RENDER_SCALE, word.Width * RENDER_SCALE, word.Height * RENDER_SCALE);
                    if (dragRect.IntersectsWith(wordRect))
                    {
                        _selectedTextData.Add(word);
                        var rect = new Rectangle { Fill = new SolidColorBrush(Color.FromArgb(100, 0, 0, 255)), Width = wordRect.Width, Height = wordRect.Height, Tag = "SELECTION" };
                        Canvas.SetLeft(rect, wordRect.X); Canvas.SetTop(rect, wordRect.Y);
                        canvas.Children.Add(rect);
                    }
                }
            }
        }

        private void ClearSelectionVisuals(Canvas canvas)
        {
            var toRemove = new List<UIElement>();
            foreach (UIElement child in canvas.Children)
                if (child is Rectangle r && r.Tag?.ToString() == "SELECTION") toRemove.Add(child);
            foreach (var r in toRemove) canvas.Children.Remove(r);
        }

        private void BtnCopyText_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedTextData.Count > 0)
            {
                var sorted = _selectedTextData.OrderBy(t => t.Y).ThenBy(t => t.X).Select(t => t.Text);
                Clipboard.SetText(string.Join(" ", sorted));
            }
            PopupTextMenu.IsOpen = false;
            RefreshPageCanvas(_focusedPage); 
        }
        private void BtnMenuHighlightY_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("HIGHLIGHT_Y");
        private void BtnMenuHighlightO_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("HIGHLIGHT_O");
        private void BtnMenuUnderline_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("UNDERLINE");

        private void ApplySelectionAnnotation(string type)
        {
            if (_selectedTextData.Count == 0) return;
            var lines = _selectedTextData.GroupBy(w => (int)(w.Y / 10));
            foreach (var line in lines)
            {
                double minX = line.Min(w => w.X);
                double minY = line.Min(w => w.Y);
                double maxX = line.Max(w => w.X + w.Width);
                double maxY = line.Max(w => w.Y + w.Height);
                var newAnn = new AnnotationData { Type = type, X = minX, Y = minY, Width = maxX - minX, Height = maxY - minY, Page = _focusedPage };
                _annotations.Add(newAnn);
            }
            RefreshPageCanvas(_focusedPage);
            PopupTextMenu.IsOpen = false;
        }

        private void DrawAnnotationsForPage(int pageIndex, Canvas canvas)
        {
            canvas.Children.Clear();
            var pageAnns = _annotations.Where(a => a.Page == pageIndex).ToList();

            foreach (var ann in pageAnns)
            {
                FrameworkElement element = null;
                double x = ann.X * RENDER_SCALE; 
                double y = ann.Y * RENDER_SCALE; 
                double w = ann.Width * RENDER_SCALE; 
                double h = ann.Height * RENDER_SCALE;

                if (ann.Type == "TEXT" || ann.Type == "OCR_TEXT")
                {
                    var tb = new TextBox
                    {
                        Text = ann.Content,
                        FontSize = ann.FontSize * RENDER_SCALE,
                        Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(ann.Color)),
                        FontFamily = new FontFamily(ann.FontFamily),
                        FontWeight = ann.IsBold ? FontWeights.Bold : FontWeights.Normal,
                        Width = w, 
                        MinHeight = Math.Max(h, 20),
                        AcceptsReturn = true, 
                        TextWrapping = TextWrapping.Wrap,
                        Background = Brushes.Transparent, 
                        BorderThickness = new Thickness(0),
                        
                        // [핵심 수정] PDF 렌더링과 일치시키기 위해 패딩 제거 및 상단 정렬
                        Padding = new Thickness(0),
                        VerticalContentAlignment = VerticalAlignment.Top
                    };
                    
                    tb.TextChanged += (s, args) => { ann.Content = tb.Text; if (tb.ActualHeight > 0) ann.Height = tb.ActualHeight / RENDER_SCALE; };
                    tb.SizeChanged += (s, args) => { if (args.NewSize.Height > 0) ann.Height = args.NewSize.Height / RENDER_SCALE; };
                    tb.GotFocus += (s, args) => { SelectAnnotation(ann); };
                    if (ann == _selectedAnnotation) { tb.Loaded += (s, e) => { tb.Focus(); tb.CaretIndex = tb.Text.Length; }; }
                    element = tb;
                }
                else 
                {
                    var rect = new Rectangle
                    {
                        Width = w, Height = h,
                        Fill = ann.Type.StartsWith("HIGHLIGHT") ? new SolidColorBrush((Color)ColorConverter.ConvertFromString(ann.Type == "HIGHLIGHT_Y" ? "#FFFF00" : "#FFA500")) { Opacity = 0.4 } : Brushes.Transparent,
                        Stroke = ann.Type == "UNDERLINE" ? Brushes.Red : Brushes.Transparent, StrokeThickness = ann.Type == "UNDERLINE" ? 2 : 0
                    };
                    rect.MouseLeftButtonDown += (s, e) => { e.Handled = true; SelectAnnotation(ann); }; rect.Cursor = Cursors.Hand; element = rect;
                }

                if (element != null)
                {
                    Canvas.SetLeft(element, x); Canvas.SetTop(element, y);
                    if (ann == _selectedAnnotation)
                    {
                        var border = new Rectangle { Width = Math.Max(w, element.ActualWidth) + 6, Height = Math.Max(h, element.ActualHeight) + 6, Stroke = Brushes.Blue, StrokeThickness = 1, StrokeDashArray = new DoubleCollection { 2 }, Fill = Brushes.Transparent, Cursor = Cursors.SizeAll };
                        Canvas.SetLeft(border, x - 3); Canvas.SetTop(border, y - 3);
                        border.MouseLeftButtonDown += (s, e) => { e.Handled = true; SelectAnnotation(ann); _isDraggingAnnot = true; _dragStartPoint = e.GetPosition(canvas); _annotStartPos = new Point(ann.X, ann.Y); _currentDrawingCanvas = canvas; canvas.CaptureMouse(); };
                        canvas.Children.Add(border); 

                        if (ann.Type == "TEXT") {
                            var handle = new Rectangle { Width = 10, Height = 10, Fill = Brushes.Red, Cursor = Cursors.SizeNWSE };
                            double actualW = Math.Max(w, element.ActualWidth); double actualH = Math.Max(h, element.ActualHeight);
                            Canvas.SetLeft(handle, x + actualW - 5); Canvas.SetTop(handle, y + actualH - 5);
                            handle.MouseLeftButtonDown += (s, e) => { e.Handled = true; _isResizingAnnot = true; _dragStartPoint = e.GetPosition(canvas); _annotStartSize = new Size(ann.Width, ann.Height); _currentDrawingCanvas = canvas; canvas.CaptureMouse(); };
                            canvas.Children.Add(handle);
                        }
                    }
                    canvas.Children.Add(element);
                }
            }
        }

        private void SelectAnnotation(AnnotationData ann)
        {
            if (_selectedAnnotation == ann) { UpdatePropertyPanel(); return; }
            var prevAnnot = _selectedAnnotation; _selectedAnnotation = ann;
            if (prevAnnot != null && prevAnnot.Page != ann.Page) RefreshPageCanvas(prevAnnot.Page);
            else if (prevAnnot != null) RefreshPageCanvas(prevAnnot.Page); 
            RefreshPageCanvas(ann.Page); UpdatePropertyPanel();
        }

        private void BtnTool_Click(object sender, RoutedEventArgs e) { if (sender is ToggleButton btn && btn.Tag != null) { string selectedTool = btn.Tag.ToString() ?? "CURSOR"; if (_currentTool == selectedTool && selectedTool != "CURSOR") SetCurrentTool("CURSOR"); else SetCurrentTool(selectedTool); } }
        private void SetCurrentTool(string tool) { _currentTool = tool; BtnCursor.IsChecked = (tool == "CURSOR"); BtnHighlightY.IsChecked = (tool == "HIGHLIGHT_Y"); BtnHighlightO.IsChecked = (tool == "HIGHLIGHT_O"); BtnUnderline.IsChecked = (tool == "UNDERLINE"); BtnTextMode.IsChecked = (tool == "TEXT"); if (_selectedAnnotation != null) { int page = _selectedAnnotation.Page; _selectedAnnotation = null; RefreshPageCanvas(page); } UpdatePropertyPanel(); }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Delete) { if (!(Keyboard.FocusedElement is TextBox)) DeleteSelectedAnnotation(); } }
        private void BtnDeleteAnnot_Click(object sender, RoutedEventArgs e) { DeleteSelectedAnnotation(); }
        private void DeleteSelectedAnnotation() { if (_selectedAnnotation == null) return; var result = MessageBox.Show("선택한 항목을 정말 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning); if (result == MessageBoxResult.Yes) { int page = _selectedAnnotation.Page; _annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; UpdatePropertyPanel(); RefreshPageCanvas(page); } }
        private void UpdatePropertyPanel() { if (_selectedAnnotation != null) { PnlProperties.Visibility = Visibility.Visible; TxtNoSelection.Visibility = Visibility.Collapsed; _isUpdatingUi = true; TxtSelectedType.Text = _selectedAnnotation.Type; TxtContent.Text = _selectedAnnotation.Content; TxtFontSize.Text = _selectedAnnotation.FontSize.ToString(); bool isText = _selectedAnnotation.Type == "TEXT"; TxtContent.IsEnabled = isText; TxtFontSize.IsEnabled = isText; CboPropColor.IsEnabled = isText; PnlTextProps.Visibility = isText ? Visibility.Visible : Visibility.Collapsed; if (isText) { bool fontFound = false; foreach(ComboBoxItem item in CboPropFontFamily.Items) { if (item.Tag?.ToString() == _selectedAnnotation.FontFamily) { item.IsSelected = true; fontFound = true; break; } } if (!fontFound && CboPropFontFamily.Items.Count > 0) CboPropFontFamily.SelectedIndex = 0; ChkPropBold.IsChecked = _selectedAnnotation.IsBold; CboPropColor.SelectedIndex = -1; foreach (ComboBoxItem item in CboPropColor.Items) { if (item.Tag?.ToString() == _selectedAnnotation.Color) { item.IsSelected = true; break; } } } _isUpdatingUi = false; } else { PnlProperties.Visibility = Visibility.Collapsed; TxtNoSelection.Visibility = Visibility.Visible; } }
        private void TxtProperty_Changed(object sender, TextChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; _selectedAnnotation.Content = TxtContent.Text; if (int.TryParse(TxtFontSize.Text, out int size)) _selectedAnnotation.FontSize = size; if (_selectedAnnotation.Type == "TEXT") { _userSettings.DefaultFontSize = _selectedAnnotation.FontSize; _userSettings.Save(); } RefreshPageCanvas(_selectedAnnotation.Page); }
        private void CboPropFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; if (CboPropFontFamily.SelectedItem is ComboBoxItem item && item.Tag != null) { _selectedAnnotation.FontFamily = item.Tag.ToString() ?? "Malgun Gothic"; _userSettings.DefaultFontFamily = _selectedAnnotation.FontFamily; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); } }
        private void CboPropColor_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; if (CboPropColor.SelectedItem is ComboBoxItem item && item.Tag != null) { _selectedAnnotation.Color = item.Tag.ToString() ?? "#FF000000"; _userSettings.DefaultColor = _selectedAnnotation.Color; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); } }
        private void ChkPropBold_Click(object sender, RoutedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; _selectedAnnotation.IsBold = ChkPropBold.IsChecked == true; _userSettings.DefaultIsBold = _selectedAnnotation.IsBold; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); }
        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if (Keyboard.Modifiers == ModifierKeys.Control) { e.Handled = true; if (e.Delta > 0) _zoomLevel += 0.1; else _zoomLevel -= 0.1; _zoomLevel = Math.Clamp(_zoomLevel, 0.2, 5.0); ApplyZoom(); } }
        private void ApplyZoom() { DocScaleTransform.ScaleX = _zoomLevel; DocScaleTransform.ScaleY = _zoomLevel; TxtZoom.Text = $"{Math.Round(_zoomLevel * 100)}%"; }
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) { if (DocumentContainer.Children.Count == 0) return; int targetIndex = Math.Clamp(_focusedPage - 1, 0, DocumentContainer.Children.Count - 1); if (DocumentContainer.Children[targetIndex] is Grid pageGrid && pageGrid.Width > 0) { double availableWidth = MainScrollViewer.ViewportWidth - 40; if (availableWidth > 0) { _zoomLevel = availableWidth / pageGrid.Width; _zoomLevel = Math.Clamp(_zoomLevel, 0.2, 5.0); ApplyZoom(); } } }
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) { if (DocumentContainer.Children.Count == 0) return; int targetIndex = Math.Clamp(_focusedPage - 1, 0, DocumentContainer.Children.Count - 1); if (DocumentContainer.Children[targetIndex] is Grid pageGrid && pageGrid.Height > 0) { double availableHeight = MainScrollViewer.ViewportHeight - 40; if (availableHeight > 0) { _zoomLevel = availableHeight / pageGrid.Height; _zoomLevel = Math.Clamp(_zoomLevel, 0.2, 5.0); ApplyZoom(); } } }
        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (_pdfData == null || _currentFilePath == null) return; try { var resultBytes = _pdfService.SavePdfWithAnnotations(_pdfData, _annotations, 1.0); File.WriteAllBytes(_currentFilePath, resultBytes); _pdfData = resultBytes; _cleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(_pdfData); MessageBox.Show("저장 완료!"); } catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); } }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) { if (_pdfData == null || _currentFilePath == null) return; var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = "Copy_" + System.IO.Path.GetFileName(_currentFilePath) }; if (dlg.ShowDialog() == true) { try { var resultBytes = _pdfService.SavePdfWithAnnotations(_pdfData, _annotations, 1.0); File.WriteAllBytes(dlg.FileName, resultBytes); _currentFilePath = dlg.FileName; _pdfData = resultBytes; _cleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(_pdfData); MessageBox.Show("저장 완료!"); } catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); } } }
        private async void BtnDeletePage_Click(object sender, RoutedEventArgs e) { int targetPage = _selectedAnnotation?.Page ?? 1; if (_totalPages <= 1 || _pdfData == null) return; _pdfData = _pdfService.DeletePage(_pdfData, targetPage - 1); _cleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(_pdfData); _annotations.RemoveAll(a => a.Page == targetPage); foreach (var ann in _annotations.Where(a => a.Page > targetPage)) ann.Page--; using (var doc = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(_pdfData), PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)) { _totalPages = doc.PageCount; } TxtPageInfo.Text = $"총 {_totalPages} 페이지"; await RenderDocument(); }
        private void BtnUndo_Click(object sender, RoutedEventArgs e) { }
    }
}