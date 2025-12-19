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
using System.Threading;       
using System.Threading.Tasks; 
using System.Collections.ObjectModel; 

namespace SmartDocProcessor.WPF
{
    public class SearchResultItem
    {
        public int Page { get; set; }
        public Rect Rect { get; set; } 
        public string Text { get; set; } = "";
    }

    public partial class MainWindow : Window
    {
        private readonly PdfService _pdfService = new PdfService();
        private readonly OcrService _ocrService = new OcrService();
        private UserSettings _userSettings;

        private ObservableCollection<PdfDocumentModel> _documents = new ObservableCollection<PdfDocumentModel>();
        private PdfDocumentModel? _activeDoc = null; 

        // [신규] 렌더링 취소 토큰
        private CancellationTokenSource? _renderCts;

        private List<SearchResultItem> _searchResults = new List<SearchResultItem>();
        private int _currentSearchIndex = -1;
        private CancellationTokenSource? _searchCts;

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
        private List<TextData> _selectedTextData = new List<TextData>();

        private const double RENDER_SCALE = 1.5; 
        private double _zoomLevel = 1.0;
        private bool _isUpdatingUi = false;

        public MainWindow()
        {
            InitializeComponent();
            _userSettings = UserSettings.Load();
            if (string.IsNullOrEmpty(_userSettings.DefaultFontFamily)) _userSettings.DefaultFontFamily = "Malgun Gothic";
            
            TabList.ItemsSource = _documents;
        }

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf", Multiselect = true };
            if (dlg.ShowDialog() == true)
            {
                foreach (var file in dlg.FileNames)
                {
                    var existing = _documents.FirstOrDefault(d => d.FilePath == file);
                    if (existing != null) { TabList.SelectedItem = existing; continue; }

                    try 
                    {
                        var data = await File.ReadAllBytesAsync(file);
                        var newDoc = new PdfDocumentModel
                        {
                            FilePath = file,
                            PdfData = data,
                            CleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(data),
                            Annotations = _pdfService.ExtractAnnotationsFromMetadata(data)
                        };

                        using (var doc = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(data), PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)) {
                            newDoc.TotalPages = doc.PageCount;
                        }

                        _documents.Add(newDoc);
                        TabList.SelectedItem = newDoc; 
                    }
                    catch (Exception ex) { MessageBox.Show($"파일 열기 실패 ({file}): {ex.Message}"); }
                }
            }
        }

        private async void TabList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is PdfDocumentModel selectedDoc)
            {
                await SwitchTab(selectedDoc);
            }
            else if (TabList.Items.Count == 0)
            {
                _activeDoc = null;
                DocumentContainer.Children.Clear();
                TxtPageInfo.Text = "문서 없음";
                CloseSearchBar();
                UpdatePropertyPanel();
            }
        }

        private async Task SwitchTab(PdfDocumentModel doc)
        {
            if (_activeDoc == doc) return;

            // [핵심] 이전 렌더링/검색 취소
            _renderCts?.Cancel();
            CancelSearch();

            _activeDoc = doc;
            _selectedAnnotation = null;
            _searchResults.Clear();
            CloseSearchBar();
            
            TxtPageInfo.Text = $"총 {_activeDoc.TotalPages} 페이지";
            TxtZoom.Text = $"{Math.Round(_activeDoc.ZoomLevel * 100)}%";
            
            // 새 렌더링 시작
            await RenderDocument();

            if (_activeDoc.PdfData != null && _pdfService.IsPdfSearchable(_activeDoc.PdfData))
            {
                await LoadTextDataForPage(_activeDoc.CurrentPage);
            }
            
            UpdatePropertyPanel();
        }

        private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PdfDocumentModel docToClose)
            {
                _documents.Remove(docToClose);
            }
        }

        // [핵심] 렌더링 취소 로직 적용
        private async System.Threading.Tasks.Task RenderDocument()
        {
            if (_activeDoc == null || (_activeDoc.CleanPdfData == null && _activeDoc.PdfData == null)) return;
            
            // 이전 작업 취소
            _renderCts?.Cancel();
            _renderCts = new CancellationTokenSource();
            var token = _renderCts.Token;

            var docToRender = _activeDoc; // 로컬 변수 캡처
            var dataToRender = docToRender.CleanPdfData ?? docToRender.PdfData;
            
            DocumentContainer.Children.Clear();
            DocScaleTransform.ScaleX = docToRender.ZoomLevel;
            DocScaleTransform.ScaleY = docToRender.ZoomLevel;

            try
            {
                for (int i = 1; i <= docToRender.TotalPages; i++)
                {
                    if (token.IsCancellationRequested) return;

                    var bitmap = await PdfRenderer.RenderPageToBitmapAsync(dataToRender!, i);
                    if (bitmap == null) continue;
                    if (token.IsCancellationRequested) return;

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
                    // 그리기 로직 호출 (이 시점에 _activeDoc가 바뀌었는지 체크)
                    if (_activeDoc == docToRender)
                    {
                        DrawAnnotationsForPage(i, canvas);
                    }
                    DocumentContainer.Children.Add(pageGrid);
                }
            }
            catch (TaskCanceledException) { }
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

        private void DrawAnnotationsForPage(int pageIndex, Canvas canvas)
        {
            if (_activeDoc == null) return;
            canvas.Children.Clear();

            // 검색 결과
            if (_searchResults.Count > 0)
            {
                var pageResults = _searchResults.Where(r => r.Page == pageIndex).ToList();
                foreach(var r in pageResults)
                {
                    bool isCurrent = (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count && _searchResults[_currentSearchIndex] == r);
                    var rect = new Rectangle
                    {
                        Width = r.Rect.Width * RENDER_SCALE,
                        Height = r.Rect.Height * RENDER_SCALE,
                        Stroke = isCurrent ? Brushes.Magenta : Brushes.Orange,
                        StrokeThickness = isCurrent ? 3 : 2,
                        Fill = isCurrent ? new SolidColorBrush(Color.FromArgb(80, 255, 0, 255)) : new SolidColorBrush(Color.FromArgb(40, 255, 165, 0))
                    };
                    Canvas.SetLeft(rect, r.Rect.X * RENDER_SCALE);
                    Canvas.SetTop(rect, r.Rect.Y * RENDER_SCALE);
                    canvas.Children.Add(rect);
                }
            }

            // 주석
            var pageAnns = _activeDoc.Annotations.Where(a => a.Page == pageIndex).ToList();
            foreach (var ann in pageAnns)
            {
                FrameworkElement element = null;
                double x = ann.X * RENDER_SCALE; double y = ann.Y * RENDER_SCALE; 
                double w = ann.Width * RENDER_SCALE; double h = ann.Height * RENDER_SCALE;

                if (ann.Type == "TEXT" || ann.Type == "OCR_TEXT")
                {
                    var tb = new TextBox
                    {
                        Text = ann.Content,
                        FontSize = ann.FontSize * RENDER_SCALE,
                        Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(ann.Color)),
                        FontFamily = new FontFamily(ann.FontFamily),
                        FontWeight = ann.IsBold ? FontWeights.Bold : FontWeights.Normal,
                        Width = w, MinHeight = Math.Max(h, 20),
                        AcceptsReturn = true, TextWrapping = TextWrapping.Wrap,
                        Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                        Padding = new Thickness(0), VerticalContentAlignment = VerticalAlignment.Top
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

        // [수정] Searchable PDF 텍스트 추출 + OCR 분기 처리
        private async Task LoadTextDataForPage(int page)
        {
            if (_activeDoc == null || _activeDoc.PdfData == null) return;
            if (!_activeDoc.PageTextData.ContainsKey(page))
            {
                // Searchable 여부 확인
                if (_pdfService.IsPdfSearchable(_activeDoc.PdfData))
                {
                    // 1. PDFSharp으로 텍스트 추출 시도
                    var texts = _pdfService.ExtractTextFromPage(_activeDoc.PdfData, page);
                    if (texts.Count > 0) 
                    {
                        _activeDoc.PageTextData[page] = texts;
                    }
                    else
                    {
                        // 2. 추출 실패 시(이미지 PDF인 경우) OCR 시도
                        var ocrTexts = await _ocrService.ExtractTextData(_activeDoc.PdfData, page);
                        if (ocrTexts.Count > 0) _activeDoc.PageTextData[page] = ocrTexts;
                    }
                }
                else
                {
                    // 3. Not Searchable -> OCR
                    var texts = await _ocrService.ExtractTextData(_activeDoc.PdfData, page);
                    if (texts.Count > 0) _activeDoc.PageTextData[page] = texts;
                }
            }
        }

        private async Task SearchPageAsync(int page, string query)
        {
            if (_activeDoc == null) return;
            await LoadTextDataForPage(page); // 통합된 메서드 호출
            SearchPageInMemory(page, query);
        }

        // --- 기타 UI 이벤트 ---
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { e.Handled = true; OpenSearchBar(); }
            if (e.Key == Key.Delete && !(Keyboard.FocusedElement is TextBox)) DeleteSelectedAnnotation();
        }
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => OpenSearchBar();
        private void OpenSearchBar() { if (_activeDoc == null) return; BdrSearchBar.Visibility = Visibility.Visible; TxtSearchQuery.Focus(); TxtSearchQuery.SelectAll(); }
        private void BtnCloseSearch_Click(object sender, RoutedEventArgs e) { CloseSearchBar(); CancelSearch(); }
        private void CloseSearchBar() { BdrSearchBar.Visibility = Visibility.Collapsed; _searchResults.Clear(); _currentSearchIndex = -1; if(_activeDoc!=null) RefreshPageCanvas(_activeDoc.CurrentPage); }
        private void CancelSearch() { _searchCts?.Cancel(); _searchCts = null; }
        private void TxtSearchQuery_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) { if (Keyboard.Modifiers == ModifierKeys.Shift) SearchPrev(); else SearchNext(); } }
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => SearchNext();
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => SearchPrev();
        private void SearchNext() { if (IsNewSearchNeeded()) { StartSearch(); return; } if (_searchResults.Count > 0) { _currentSearchIndex++; if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0; HighlightCurrentSearchResult(); } }
        private void SearchPrev() { if (IsNewSearchNeeded()) { StartSearch(); return; } if (_searchResults.Count > 0) { _currentSearchIndex--; if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1; HighlightCurrentSearchResult(); } }
        private bool IsNewSearchNeeded() { string currentQuery = TxtSearchQuery.Text.Trim(); if (_searchResults.Count == 0 && _searchCts == null) return true; if (_searchResults.Count > 0 && _searchResults[0].Text != currentQuery) return true; return false; }
        
        private async void StartSearch()
        {
            if (_activeDoc == null) return; string query = TxtSearchQuery.Text.Trim(); if (string.IsNullOrEmpty(query)) return;
            CancelSearch(); _searchCts = new CancellationTokenSource(); var token = _searchCts.Token;
            _searchResults.Clear(); _currentSearchIndex = -1; TxtSearchCount.Text = "검색중..";
            
            await SearchPageAsync(_activeDoc.CurrentPage, query);
            if (_searchResults.Count > 0) { _currentSearchIndex = 0; HighlightCurrentSearchResult(); }

            var activeDocRef = _activeDoc; // 검색 도중 탭 전환 대비 참조 캡처
            _ = Task.Run(async () => {
                var pageOrder = Enumerable.Range(activeDocRef.CurrentPage + 1, activeDocRef.TotalPages - activeDocRef.CurrentPage).Concat(Enumerable.Range(1, activeDocRef.CurrentPage - 1));
                foreach (int page in pageOrder) {
                    if (token.IsCancellationRequested) break;
                    // 백그라운드에서는 텍스트 로드만 수행 (SearchPageAsync 로직의 일부)
                    bool hasData = activeDocRef.PageTextData.ContainsKey(page);
                    if (!hasData) {
                        // 백그라운드 스레드에서 Extract 호출 불가할 수 있으므로 UI 스레드 사용
                        await Dispatcher.InvokeAsync(async () => {
                            if (activeDocRef == _activeDoc) await LoadTextDataForPage(page); 
                        });
                    }
                    await Dispatcher.InvokeAsync(() => {
                        if (token.IsCancellationRequested || activeDocRef != _activeDoc) return; // 탭 바뀌었으면 중단
                        SearchPageInMemory(page, query);
                        if (_searchResults.Count > 0 && _currentSearchIndex == -1) { _currentSearchIndex = 0; HighlightCurrentSearchResult(); } else UpdateSearchCountText();
                    });
                }
                await Dispatcher.InvokeAsync(() => { if (!token.IsCancellationRequested) { if (_searchResults.Count == 0) MessageBox.Show("검색 결과가 없습니다."); else UpdateSearchCountText(); } });
            }, token);
        }

        private void SearchPageInMemory(int page, string query) { if (_activeDoc == null || !_activeDoc.PageTextData.ContainsKey(page)) return; bool foundAny = false; foreach (var item in _activeDoc.PageTextData[page]) { if (item.Text.Contains(query, StringComparison.OrdinalIgnoreCase)) { _searchResults.Add(new SearchResultItem { Page = page, Rect = new Rect(item.X, item.Y, item.Width, item.Height), Text = query }); foundAny = true; } } if (foundAny) _searchResults = _searchResults.OrderBy(r => r.Page).ThenBy(r => r.Rect.Y).ToList(); }
        private void HighlightCurrentSearchResult() { if (_searchResults.Count == 0) return; if (_currentSearchIndex < 0) _currentSearchIndex = 0; if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0; var result = _searchResults[_currentSearchIndex]; UpdateSearchCountText(); if (_activeDoc != null && _activeDoc.CurrentPage != result.Page) { _activeDoc.CurrentPage = result.Page; if (DocumentContainer.Children.Count >= _activeDoc.CurrentPage) { var pageGrid = DocumentContainer.Children[_activeDoc.CurrentPage - 1] as FrameworkElement; pageGrid?.BringIntoView(); } } RefreshPageCanvas(_activeDoc?.CurrentPage ?? 1); }
        private void UpdateSearchCountText() => TxtSearchCount.Text = _searchResults.Count > 0 ? $"{_currentSearchIndex + 1}/{_searchResults.Count}" : "0/0";

        private async void BtnOcr_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDoc == null || _activeDoc.PdfData == null) return;
            MessageBox.Show("전체 페이지 OCR 분석을 시작합니다. 완료되면 드래그가 가능해집니다.");
            this.Cursor = Cursors.Wait;
            for (int i = 1; i <= _activeDoc.TotalPages; i++) {
                if (!_activeDoc.PageTextData.ContainsKey(i)) {
                    var texts = await _ocrService.ExtractTextData(_activeDoc.PdfData, i);
                    if (texts.Count > 0) _activeDoc.PageTextData[i] = texts;
                }
            }
            this.Cursor = Cursors.Arrow;
            MessageBox.Show("OCR 완료! 이제 텍스트 드래그가 가능합니다.");
        }

        private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_activeDoc == null || sender is not Canvas canvas) return;
            if (canvas.Tag is int pageNum) { 
                _activeDoc.CurrentPage = pageNum;
                // 클릭 시 데이터 로드 (드래그 준비)
                if (!_activeDoc.PageTextData.ContainsKey(pageNum)) { _ = LoadTextDataForPage(pageNum); }
            }
            _dragStartPoint = e.GetPosition(canvas);
            _currentDrawingCanvas = canvas;
            canvas.CaptureMouse();

            if (_currentTool == "CURSOR" && _activeDoc.PageTextData.ContainsKey(_activeDoc.CurrentPage))
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
            if (_currentDrawingCanvas == null || _activeDoc == null) return;
            var currentPoint = e.GetPosition(_currentDrawingCanvas);
            if (_isTextSelecting) { UpdateTextSelection(_currentDrawingCanvas, _dragStartPoint, currentPoint); return; }
            if (_isDraggingAnnot && _selectedAnnotation != null) { double deltaX = (currentPoint.X - _dragStartPoint.X) / RENDER_SCALE; double deltaY = (currentPoint.Y - _dragStartPoint.Y) / RENDER_SCALE; _selectedAnnotation.X = _annotStartPos.X + deltaX; _selectedAnnotation.Y = _annotStartPos.Y + deltaY; RefreshPageCanvas(_selectedAnnotation.Page); } 
            else if (_isResizingAnnot && _selectedAnnotation != null) { double deltaX = (currentPoint.X - _dragStartPoint.X) / RENDER_SCALE; double deltaY = (currentPoint.Y - _dragStartPoint.Y) / RENDER_SCALE; double newW = Math.Max(10, _annotStartSize.Width + deltaX); double newH = Math.Max(10, _annotStartSize.Height + deltaY); _selectedAnnotation.Width = newW; _selectedAnnotation.Height = newH; RefreshPageCanvas(_selectedAnnotation.Page); }
        }

        private void Canvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            _currentDrawingCanvas?.ReleaseMouseCapture();
            if (_activeDoc == null) return;
            if (_isTextSelecting) { _isTextSelecting = false; if (_selectedTextData.Count > 0) PopupTextMenu.IsOpen = true; _currentDrawingCanvas = null; return; }
            if (_isDraggingAnnot || _isResizingAnnot) { _isDraggingAnnot = false; _isResizingAnnot = false; _currentDrawingCanvas = null; UpdatePropertyPanel(); return; }
            if (!_isDrawing || _currentDrawingCanvas == null) return;
            _isDrawing = false; var canvas = _currentDrawingCanvas; var endPoint = e.GetPosition(canvas); _currentDrawingCanvas = null; double x = Math.Min(_dragStartPoint.X, endPoint.X); double y = Math.Min(_dragStartPoint.Y, endPoint.Y); double w = Math.Abs(endPoint.X - _dragStartPoint.X); double h = Math.Abs(endPoint.Y - _dragStartPoint.Y); if (w < 5 && h < 5 && _currentTool != "TEXT") return; int pageNum = (canvas.Tag as int?) ?? 1; var newAnn = new AnnotationData { Type = _currentTool, X = x / RENDER_SCALE, Y = y / RENDER_SCALE, Width = Math.Max(w / RENDER_SCALE, 100), Height = Math.Max(h / RENDER_SCALE, 30), Page = pageNum }; if (_currentTool == "TEXT") { newAnn.Content = "텍스트 입력"; newAnn.FontFamily = _userSettings.DefaultFontFamily; newAnn.FontSize = _userSettings.DefaultFontSize; newAnn.IsBold = _userSettings.DefaultIsBold; newAnn.Color = _userSettings.DefaultColor; } _activeDoc.Annotations.Add(newAnn); if (_currentTool == "TEXT") SelectAnnotation(newAnn); else RefreshPageCanvas(pageNum);
        }

        private void UpdateTextSelection(Canvas canvas, Point start, Point end) { if (_activeDoc == null) return; double x = Math.Min(start.X, end.X); double y = Math.Min(start.Y, end.Y); double w = Math.Abs(end.X - start.X); double h = Math.Abs(end.Y - start.Y); var dragRect = new Rect(x, y, w, h); ClearSelectionVisuals(canvas); _selectedTextData.Clear(); if (_activeDoc.PageTextData.ContainsKey(_activeDoc.CurrentPage)) { foreach (var word in _activeDoc.PageTextData[_activeDoc.CurrentPage]) { Rect wordRect = new Rect(word.X * RENDER_SCALE, word.Y * RENDER_SCALE, word.Width * RENDER_SCALE, word.Height * RENDER_SCALE); if (dragRect.IntersectsWith(wordRect)) { _selectedTextData.Add(word); var rect = new Rectangle { Fill = new SolidColorBrush(Color.FromArgb(100, 0, 0, 255)), Width = wordRect.Width, Height = wordRect.Height, Tag = "SELECTION" }; Canvas.SetLeft(rect, wordRect.X); Canvas.SetTop(rect, wordRect.Y); canvas.Children.Add(rect); } } } }
        private void ClearSelectionVisuals(Canvas canvas) { var toRemove = new List<UIElement>(); foreach (UIElement child in canvas.Children) if (child is Rectangle r && r.Tag?.ToString() == "SELECTION") toRemove.Add(child); foreach (var r in toRemove) canvas.Children.Remove(r); }
        private void BtnCopyText_Click(object sender, RoutedEventArgs e) { if (_selectedTextData.Count > 0) { var sorted = _selectedTextData.OrderBy(t => t.Y).ThenBy(t => t.X).Select(t => t.Text); Clipboard.SetText(string.Join(" ", sorted)); } PopupTextMenu.IsOpen = false; if(_activeDoc!=null) RefreshPageCanvas(_activeDoc.CurrentPage); }
        private void BtnMenuHighlightY_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("HIGHLIGHT_Y");
        private void BtnMenuHighlightO_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("HIGHLIGHT_O");
        private void BtnMenuUnderline_Click(object sender, RoutedEventArgs e) => ApplySelectionAnnotation("UNDERLINE");
        private void ApplySelectionAnnotation(string type) { if (_selectedTextData.Count == 0 || _activeDoc == null) return; var lines = _selectedTextData.GroupBy(w => (int)(w.Y / 10)); foreach (var line in lines) { double minX = line.Min(w => w.X); double minY = line.Min(w => w.Y); double maxX = line.Max(w => w.X + w.Width); double maxY = line.Max(w => w.Y + w.Height); var newAnn = new AnnotationData { Type = type, X = minX, Y = minY, Width = maxX - minX, Height = maxY - minY, Page = _activeDoc.CurrentPage }; _activeDoc.Annotations.Add(newAnn); } RefreshPageCanvas(_activeDoc.CurrentPage); PopupTextMenu.IsOpen = false; }
        private void SelectAnnotation(AnnotationData ann) { if (_selectedAnnotation == ann) { UpdatePropertyPanel(); return; } var prevAnnot = _selectedAnnotation; _selectedAnnotation = ann; if (prevAnnot != null && prevAnnot.Page != ann.Page) RefreshPageCanvas(prevAnnot.Page); else if (prevAnnot != null) RefreshPageCanvas(prevAnnot.Page); RefreshPageCanvas(ann.Page); UpdatePropertyPanel(); }
        private void BtnTool_Click(object sender, RoutedEventArgs e) { if (sender is ToggleButton btn && btn.Tag != null) { string selectedTool = btn.Tag.ToString() ?? "CURSOR"; if (_currentTool == selectedTool && selectedTool != "CURSOR") SetCurrentTool("CURSOR"); else SetCurrentTool(selectedTool); } }
        private void SetCurrentTool(string tool) { _currentTool = tool; BtnCursor.IsChecked = (tool == "CURSOR"); BtnHighlightY.IsChecked = (tool == "HIGHLIGHT_Y"); BtnHighlightO.IsChecked = (tool == "HIGHLIGHT_O"); BtnUnderline.IsChecked = (tool == "UNDERLINE"); BtnTextMode.IsChecked = (tool == "TEXT"); if (_selectedAnnotation != null) { int page = _selectedAnnotation.Page; _selectedAnnotation = null; RefreshPageCanvas(page); } UpdatePropertyPanel(); }
        private void BtnDeleteAnnot_Click(object sender, RoutedEventArgs e) { DeleteSelectedAnnotation(); }
        private void DeleteSelectedAnnotation() { if (_selectedAnnotation == null || _activeDoc == null) return; var result = MessageBox.Show("선택한 항목을 정말 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning); if (result == MessageBoxResult.Yes) { int page = _selectedAnnotation.Page; _activeDoc.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; UpdatePropertyPanel(); RefreshPageCanvas(page); } }
        private void UpdatePropertyPanel() { if (_selectedAnnotation != null) { PnlProperties.Visibility = Visibility.Visible; TxtNoSelection.Visibility = Visibility.Collapsed; _isUpdatingUi = true; TxtSelectedType.Text = _selectedAnnotation.Type; TxtContent.Text = _selectedAnnotation.Content; TxtFontSize.Text = _selectedAnnotation.FontSize.ToString(); bool isText = _selectedAnnotation.Type == "TEXT"; TxtContent.IsEnabled = isText; TxtFontSize.IsEnabled = isText; CboPropColor.IsEnabled = isText; PnlTextProps.Visibility = isText ? Visibility.Visible : Visibility.Collapsed; if (isText) { bool fontFound = false; foreach(ComboBoxItem item in CboPropFontFamily.Items) { if (item.Tag?.ToString() == _selectedAnnotation.FontFamily) { item.IsSelected = true; fontFound = true; break; } } if (!fontFound && CboPropFontFamily.Items.Count > 0) CboPropFontFamily.SelectedIndex = 0; ChkPropBold.IsChecked = _selectedAnnotation.IsBold; CboPropColor.SelectedIndex = -1; foreach (ComboBoxItem item in CboPropColor.Items) { if (item.Tag?.ToString() == _selectedAnnotation.Color) { item.IsSelected = true; break; } } } _isUpdatingUi = false; } else { PnlProperties.Visibility = Visibility.Collapsed; TxtNoSelection.Visibility = Visibility.Visible; } }
        private void TxtProperty_Changed(object sender, TextChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; _selectedAnnotation.Content = TxtContent.Text; if (int.TryParse(TxtFontSize.Text, out int size)) _selectedAnnotation.FontSize = size; if (_selectedAnnotation.Type == "TEXT") { _userSettings.DefaultFontSize = _selectedAnnotation.FontSize; _userSettings.Save(); } RefreshPageCanvas(_selectedAnnotation.Page); }
        private void CboPropFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; if (CboPropFontFamily.SelectedItem is ComboBoxItem item && item.Tag != null) { _selectedAnnotation.FontFamily = item.Tag.ToString() ?? "Malgun Gothic"; _userSettings.DefaultFontFamily = _selectedAnnotation.FontFamily; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); } }
        private void CboPropColor_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; if (CboPropColor.SelectedItem is ComboBoxItem item && item.Tag != null) { _selectedAnnotation.Color = item.Tag.ToString() ?? "#FF000000"; _userSettings.DefaultColor = _selectedAnnotation.Color; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); } }
        private void ChkPropBold_Click(object sender, RoutedEventArgs e) { if (_selectedAnnotation == null || _isUpdatingUi) return; _selectedAnnotation.IsBold = ChkPropBold.IsChecked == true; _userSettings.DefaultIsBold = _selectedAnnotation.IsBold; _userSettings.Save(); RefreshPageCanvas(_selectedAnnotation.Page); }
        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if (Keyboard.Modifiers == ModifierKeys.Control && _activeDoc != null) { e.Handled = true; if (e.Delta > 0) _activeDoc.ZoomLevel += 0.1; else _activeDoc.ZoomLevel -= 0.1; _activeDoc.ZoomLevel = Math.Clamp(_activeDoc.ZoomLevel, 0.2, 5.0); ApplyZoom(); } }
        private void ApplyZoom() { if(_activeDoc == null) return; DocScaleTransform.ScaleX = _activeDoc.ZoomLevel; DocScaleTransform.ScaleY = _activeDoc.ZoomLevel; TxtZoom.Text = $"{Math.Round(_activeDoc.ZoomLevel * 100)}%"; }
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) { if (DocumentContainer.Children.Count == 0 || _activeDoc == null) return; int targetIndex = Math.Clamp(_activeDoc.CurrentPage - 1, 0, DocumentContainer.Children.Count - 1); if (DocumentContainer.Children[targetIndex] is Grid pageGrid && pageGrid.Width > 0) { double availableWidth = MainScrollViewer.ViewportWidth - 40; if (availableWidth > 0) { _activeDoc.ZoomLevel = availableWidth / pageGrid.Width; _activeDoc.ZoomLevel = Math.Clamp(_activeDoc.ZoomLevel, 0.2, 5.0); ApplyZoom(); } } }
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) { if (DocumentContainer.Children.Count == 0 || _activeDoc == null) return; int targetIndex = Math.Clamp(_activeDoc.CurrentPage - 1, 0, DocumentContainer.Children.Count - 1); if (DocumentContainer.Children[targetIndex] is Grid pageGrid && pageGrid.Height > 0) { double availableHeight = MainScrollViewer.ViewportHeight - 40; if (availableHeight > 0) { _activeDoc.ZoomLevel = availableHeight / pageGrid.Height; _activeDoc.ZoomLevel = Math.Clamp(_activeDoc.ZoomLevel, 0.2, 5.0); ApplyZoom(); } } }
        private void BtnUndo_Click(object sender, RoutedEventArgs e) { }

        private List<AnnotationData> PrepareAnnotationsForSave()
        {
            if (_activeDoc == null) return new List<AnnotationData>();
            var allAnns = new List<AnnotationData>(_activeDoc.Annotations);
            foreach (var kvp in _activeDoc.PageTextData)
            {
                int pageNum = kvp.Key;
                foreach (var textData in kvp.Value)
                {
                    allAnns.Add(new AnnotationData { Type = "OCR_TEXT", Content = textData.Text, X = textData.X, Y = textData.Y, Width = textData.Width, Height = textData.Height, Page = pageNum, FontSize = 10 });
                }
            }
            return allAnns;
        }

        private async void BtnDeletePage_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDoc == null) return;
            int targetPage = _selectedAnnotation?.Page ?? _activeDoc.CurrentPage;
            if (_activeDoc.TotalPages <= 1 || _activeDoc.PdfData == null) return;
            _activeDoc.PdfData = _pdfService.DeletePage(_activeDoc.PdfData, targetPage - 1);
            _activeDoc.CleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(_activeDoc.PdfData);
            _activeDoc.Annotations.RemoveAll(a => a.Page == targetPage);
            foreach (var ann in _activeDoc.Annotations.Where(a => a.Page > targetPage)) ann.Page--;
            using (var doc = PdfSharp.Pdf.IO.PdfReader.Open(new MemoryStream(_activeDoc.PdfData), PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)) { _activeDoc.TotalPages = doc.PageCount; }
            TxtPageInfo.Text = $"총 {_activeDoc.TotalPages} 페이지";
            await RenderDocument();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDoc == null || _activeDoc.PdfData == null) return;
            try
            {
                var allAnns = PrepareAnnotationsForSave();
                var resultBytes = _pdfService.SavePdfWithAnnotations(_activeDoc.PdfData, allAnns, _activeDoc.ZoomLevel);
                File.WriteAllBytes(_activeDoc.FilePath, resultBytes);
                _activeDoc.PdfData = resultBytes;
                _activeDoc.CleanPdfData = _pdfService.GetPdfBytesWithoutAnnotations(resultBytes);
                MessageBox.Show("저장 완료!");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }

        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (_activeDoc == null || _activeDoc.PdfData == null) return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = "Copy_" + _activeDoc.FileName };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var allAnns = PrepareAnnotationsForSave();
                    var resultBytes = _pdfService.SavePdfWithAnnotations(_activeDoc.PdfData, allAnns, 1.0);
                    File.WriteAllBytes(dlg.FileName, resultBytes);
                    MessageBox.Show("저장 완료!");
                }
                catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
            }
        }
    }
}