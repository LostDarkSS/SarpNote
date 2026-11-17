using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using PdfiumViewer;
using GdiBitmap = System.Drawing.Bitmap;
using GdiColor = System.Drawing.Color;
using GdiGraphics = System.Drawing.Graphics;
using GdiPen = System.Drawing.Pen;
using GdiSolidBrush = System.Drawing.SolidBrush;
using IOPath = System.IO.Path;

namespace SarpNote
{
    public class AppMetadata
    {
        public Dictionary<string, int> LastPages { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, string> OriginalPaths { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> Tags { get; set; } = new Dictionary<string, string>();
        public List<string> Bookmarks { get; set; } = new List<string>();
        public int HoldToLineMs { get; set; } = 550;
        public string Language { get; set; } = "tr";
        public double StylusSmoothing { get; set; } = 0.8;
        public int BarrelButtonMode { get; set; } = 0;
        public bool EraserAuraEnabled { get; set; } = true;
        public double RenderScale { get; set; } = 2.1;
        public int MaxCachePages { get; set; } = 24;
        public double ScrollSpeed { get; set; } = 1.05;
        public double DefaultZoom { get; set; } = 1.30;
        public int PageFilterMode { get; set; } = 0;
        public int MarginWidthPx { get; set; } = 560;
        public string MarginColorHex { get; set; } = "#FAFAFC";
        public bool StartWithDrawerOpen { get; set; } = true;
        public int StartupBehavior { get; set; } = 0; // 0: Son PDF, 1: Kitaplık
        public string StartupPenStyle { get; set; } = "Pro";
        public string StartupPenColorHex { get; set; } = "#0F172A";
        public double StartupPenThickness { get; set; } = 1.5;
        public double StartupEraserSize { get; set; } = 10.0;
        public bool StartupExtraMargin { get; set; } = false;
        public bool StartMaximized { get; set; } = true;
    }

    public class PageThumbnailItem : INotifyPropertyChanged
    {
        private ImageSource _thumbnailSource;
        private Visibility _starVisibility = Visibility.Collapsed;
        public int PageIndex { get; set; }
        public string PageNumberText => (PageIndex + 1).ToString();
        public ImageSource ThumbnailSource
        {
            get => _thumbnailSource;
            set { if (_thumbnailSource != value) { _thumbnailSource = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ThumbnailSource))); } }
        }
        public Visibility StarVisibility
        {
            get => _starVisibility;
            set { if (_starVisibility != value) { _starVisibility = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StarVisibility))); } }
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class PageSlot
    {
        public int PageIndex { get; set; }
        public double PdfWidth { get; set; }
        public double PdfHeight { get; set; }
        public int ExtraMarginWidth { get; set; }
        public double TotalWidth => PdfWidth + ExtraMarginWidth;
        public double TopOffset { get; set; }
        public Grid Container { get; set; }
        public Image PdfImage { get; set; }
        public Image MarginImage { get; set; }
        public Canvas EraserHoverCanvas { get; set; }
        
        public InkCanvas Canvas { get; set; }
        public bool IsRealized { get; set; }
        public bool HasHighResRendered { get; set; }
        public byte[] CachedStrokeBytes { get; set; }
        public List<byte[]> History { get; } = new List<byte[]>();
        public int HistoryIndex { get; set; } = -1;
        public bool HasUncommittedChanges { get; set; }
    }

    // DRAWBOARD TARZI COKLU SEKME BELLEK OTURUMU (SIFIR GECIKMELI SEKME GECISI)
    public class PdfTabSession
    {
        public string FileName { get; set; }
        public string FullPath { get; set; }
        public string PdfKey { get; set; }
        public PdfDocument Document { get; set; }
        public readonly object SyncLock = new object();
        public List<PageSlot> Slots { get; } = new List<PageSlot>();
        public List<PageThumbnailItem> AllThumbnails { get; } = new List<PageThumbnailItem>();
        public Dictionary<int, BitmapSource> LowResCache { get; } = new Dictionary<int, BitmapSource>();
        public Dictionary<int, BitmapSource> HighResCache { get; } = new Dictionary<int, BitmapSource>();
        public LinkedList<int> HighResLru { get; } = new LinkedList<int>();
        public int CurrentPage { get; set; }
        public double ZoomLevel { get; set; } = 1.30;
        public double ScrollOffsetX { get; set; }
        public double ScrollOffsetY { get; set; }
        public bool HasInitialScrollSet { get; set; }
        public bool ExtraMarginEnabled { get; set; }
        public int PaperStyleIndex { get; set; }

        public void CacheHighRes(int pageIndex, BitmapSource bmp, int maxPages = 18)
        {
            HighResCache[pageIndex] = bmp;
            HighResLru.Remove(pageIndex);
            HighResLru.AddLast(pageIndex);
            while (HighResLru.Count > maxPages)
            {
                int oldest = HighResLru.First.Value;
                HighResLru.RemoveFirst();
                HighResCache.Remove(oldest);
            }
        }
    }

    internal sealed class SafeCursorHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);
        public SafeCursorHandle(IntPtr preexistingHandle) : base(true) { SetHandle(preexistingHandle); }
        protected override bool ReleaseHandle() => DestroyIcon(handle);
    }

    public partial class MainWindow : Window
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct ICONINFO { public bool fIcon; public int xHotspot; public int yHotspot; public IntPtr hbmMask; public IntPtr hbmColor; }
        [DllImport("user32.dll")] private static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);
        [DllImport("user32.dll")] private static extern IntPtr CreateIconIndirect([In] ref ICONINFO piconinfo);
        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr hIcon);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr hObject);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int HTCAPTION = 0x0002;

        private Cursor _hardwareCursor;
        private readonly Dictionary<string, Cursor> _cursorCache = new Dictionary<string, Cursor>();
        private readonly string _baseDir, _libraryFolder, _annotationsFolder, _coversFolder, _metaFilePath;
        private AppMetadata _meta = new AppMetadata();

        // Sekme Oturum Yöneticisi
        private readonly Dictionary<string, PdfTabSession> _tabSessions = new Dictionary<string, PdfTabSession>(StringComparer.OrdinalIgnoreCase);
        private PdfTabSession _activeSession;
        private PdfDocument _currentPdf => _activeSession?.Document;
        private string _currentPdfFileName => _activeSession?.FileName;
        private string _currentPdfKey => _activeSession?.PdfKey;
        private int _currentPage
        {
            get => _activeSession?.CurrentPage ?? 0;
            set { if (_activeSession != null) _activeSession.CurrentPage = value; }
        }

        private readonly List<string> _openTabs = new List<string>();
        private bool _isPagesDrawerOpen = true, _filterStarredPagesOnly = false, _dashboardDirty = true;
        private string _dashboardCategory = "All";
        private readonly ConcurrentDictionary<string, ImageSource> _coverCache = new ConcurrentDictionary<string, ImageSource>(StringComparer.OrdinalIgnoreCase);

        private readonly ObservableCollection<PageThumbnailItem> _thumbnailItems = new ObservableCollection<PageThumbnailItem>();
        private List<PageThumbnailItem> _allThumbnailItems => _activeSession?.AllThumbnails ?? _emptyThumbnails;
        private readonly List<PageThumbnailItem> _emptyThumbnails = new List<PageThumbnailItem>();
        private bool _suppressThumbnailSelection = false;

        private List<PageSlot> _pageSlots => _activeSession?.Slots ?? _emptySlots;
        private readonly List<PageSlot> _emptySlots = new List<PageSlot>();
        private PageSlot _activeSlot;
        private bool _suppressScrollSync = false, _suppressHistory = false, _suppressStudioEvents = false, _suppressMarginEvent = false;

        // Arka Plan Asenkron Render Motoru
        private CancellationTokenSource _renderCts = new CancellationTokenSource();
        private volatile int _visibleCenterPage = 0;

        private readonly Dictionary<string, double> _toolSizes = new Dictionary<string, double> { ["Pen"] = 1.5, ["Highlighter"] = 20.0, ["Eraser"] = 10.0 };
        private Color _penColor = (Color)ColorConverter.ConvertFromString("#0F172A");
        private Color _highlighterColor = (Color)ColorConverter.ConvertFromString("#FACC15");
        private Color _activeColor = (Color)ColorConverter.ConvertFromString("#0F172A");
        private string _activeTool = "Pen", _eraserSubMode = "Stroke", _penStyleMode = "Pro";
        private bool _tempEraserFromButton = false;
        private string _previousToolBeforeTempEraser = "Pen";
        private int _activeShapePreset = 0;
        private double _penThickness = 2.0, _penOpacity = 1.0, _highlighterOpacity = 0.88;
        private double _pickerHue = 222.0, _pickerSat = 0.64, _pickerVal = 0.16;
        private bool _isDraggingSatVal = false, _isDraggingHue = false, _suppressHexEvent = false;

        // V-Sync Ultra-Smooth Zoom & Kaydırma Motoru
        private double _zoomLevel = 1.30, _targetZoom = 1.30;
        private Point _zoomAnchorViewport, _zoomAnchorDoc;
        private double _smoothCurrentY = 0, _smoothTargetY = 0;
        private bool _isSmoothZooming = false, _isSmoothScrolling = false, _renderLoopAttached = false;
        private readonly DispatcherTimer _zoomSettleTimer;
        private readonly DispatcherTimer _thumbRenderDebounceTimer;

        private bool _isPanning = false;
        private Point _panStartMouse, _panStartScrollOffset;
        private double _panVelocityY = 0, _lastPanY = 0;
        private DateTime _lastPanTime = DateTime.UtcNow;

        private readonly DispatcherTimer _holdTimer;
        private readonly List<Point> _rawPoints = new List<Point>();
        private Point _dragStartPoint, _lastMovePoint;
        private bool _isPointerDown = false, _shapeConvertedLive = false;
        private Stroke _previewShapeStroke = null;
        private Stroke _activeFreehandStroke = null;
        private StylusPointCollection _activeFreehandPoints = null;
        private RenderTargetBitmap _cachedMarginBitmap;
        private int _cachedMarginStyle = -1, _cachedMarginHeight = -1;

        private static readonly string[] AvailableTags = { "Matematik", "Fizik", "Kimya", "Biyoloji", "Deneme" };

        public MainWindow()
        {
            InitializeComponent();
            if (Application.Current != null) Application.Current.DispatcherUnhandledException += (s, exArgs) => { exArgs.Handled = true; };
            TryLoadSafeWindowIcon();
            MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight;
            MaxWidth = SystemParameters.MaximizedPrimaryScreenWidth;
            ThumbnailList.ItemsSource = _thumbnailItems;
            this.PreviewStylusButtonDown += (s, e) =>
        {
            if ((_meta.BarrelButtonMode == 0 && e.StylusButton.Guid == StylusPointProperties.BarrelButton.Id) ||
                (e.StylusButton.Guid == StylusPointProperties.TipButton.Id && e.StylusDevice.Inverted))
            {
                if (_meta.BarrelButtonMode == 1 && e.StylusButton.Guid == StylusPointProperties.BarrelButton.Id)
                {
                    string nextTool = _activeTool == "Eraser" ? "Pen" : "Eraser";
                    Tool_Click(nextTool == "Eraser" ? BtnEraser : BtnPen, new RoutedEventArgs());
                }
                else
                {
                    ActivateTempEraser();
                }
                e.Handled = true;
            }
        };
        this.PreviewStylusButtonUp += (s, e) =>
        {
            if (e.StylusButton.Guid == StylusPointProperties.BarrelButton.Id && _tempEraserFromButton)
            {
                DeactivateTempEraser();
                e.Handled = true;
            }
        };
        this.PreviewStylusInAirMove += (s, e) =>
        {
            if (e.StylusDevice == null) return;
            if (_meta.BarrelButtonMode == 0)
            {
                bool barrelDown = e.StylusDevice.Inverted || (_meta.BarrelButtonMode == 0 && e.StylusDevice.StylusButtons.Any(b => b.Guid == StylusPointProperties.BarrelButton.Id && b.StylusButtonState == StylusButtonState.Down));
                if (barrelDown && !_tempEraserFromButton) ActivateTempEraser();
                else if (!barrelDown && _tempEraserFromButton) DeactivateTempEraser();
            }

            if (_activeTool == "Eraser" && DashboardOverlay.Visibility != Visibility.Visible)
            {
                PageSlot hs = FindSlotAtStackPoint(e.GetPosition(PagesStack));
                if (hs != null && hs.IsRealized && hs.Canvas != null)
                    RenderEraserHover(hs, e.GetPosition(hs.Canvas));
                else
                    ClearEraserHover();
            }
        };
        this.PreviewMouseRightButtonDown += (s, e) =>
            {
                if (DashboardOverlay.Visibility == Visibility.Visible) return;
                ActivateTempEraser();
                e.Handled = true;
            };
            this.PreviewMouseRightButtonUp += (s, e) =>
            {
                if (_tempEraserFromButton)
                {
                    _isPointerDown = false;
                    DeactivateTempEraser();
                    e.Handled = true;
                }
            };

            _baseDir = IOPath.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SarpNote_Library");
            _libraryFolder = IOPath.Combine(_baseDir, "PDFs");
            _annotationsFolder = IOPath.Combine(_baseDir, "Annotations");
            _coversFolder = IOPath.Combine(_baseDir, "Covers");
            _metaFilePath = IOPath.Combine(_baseDir, "metadata.json");
            Directory.CreateDirectory(_libraryFolder);
            Directory.CreateDirectory(_annotationsFolder);
            Directory.CreateDirectory(_coversFolder);
            LoadMetadata();

            _holdTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(550) };
            _holdTimer.Tick += HoldTimer_Tick;

            _zoomSettleTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(110) };
            _zoomSettleTimer.Tick += (s, e) =>
            {
                _zoomSettleTimer.Stop();
            if (_activeTool != "Eraser") ClearEraserHover();
            RebuildHardwareCursor();
                ApplyToolSettingsToAllRealizedPages();
                UpdateVisiblePagesVirtualization();
                TriggerBackgroundPageRender();
            };

            _thumbRenderDebounceTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(35) };
            _thumbRenderDebounceTimer.Tick += (s, e) =>
            {
                _thumbRenderDebounceTimer.Stop();
                TriggerBackgroundThumbnailRender();
            };

            Loaded += async (s, e) => await RunStartupSequenceAsync();
        }

        private void SetStartupProgress(string status, double ratio)
        {
            double clamped = Math.Clamp(ratio, 0.08, 1.0);
            if (TxtStartupStatus != null) TxtStartupStatus.Text = status;
            if (TxtStartupPercent != null) TxtStartupPercent.Text = $"%{(int)Math.Round(clamped * 100)}";
            if (StartupProgressFill != null)
            {
                double targetW = clamped * 330.0;
                var anim = new System.Windows.Media.Animation.DoubleAnimation(targetW, TimeSpan.FromMilliseconds(160))
                {
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
                };
                StartupProgressFill.BeginAnimation(FrameworkElement.WidthProperty, anim);
            }
        }

        private async Task RunStartupSequenceAsync()
        {
            _currentLang = LoadPersistedLanguageCode();
            UpdateWindowMaxRestoreIcon();
            ApplyMaximizedBoundsFix();
            SetStartupProgress("Arayüz ve çizim motoru başlatılıyor...", 0.15);

            // Pencerenin ve yükleme ekranının anında (0 ms) ekrana çizilmesini sağla
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            await Task.Delay(25);

            EnsureProFeaturesInitialized();
            EnsureLanguageSelectorInSettings();
            _currentLang = LoadPersistedLanguageCode(); ApplyApplicationLanguage(_currentLang, saveToMeta: false);
            RebuildHardwareCursor();
            UpdatePenStudioVisuals();
            HighlightActiveToolButtons(); ApplyLoadedSettingsToRuntime(); SetStartupProgress("Kütüphane ve son oturum taranıyor...", 0.35);

            string firstPdf = null;
            int targetPage = 0;

            await Task.Run(() =>
            {
                RestoreBooksIfWiped(); Dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(_meta.StartupPenColorHex))
                            _penColor = _activeColor = (Color)ColorConverter.ConvertFromString(_meta.StartupPenColorHex);
                    }
                    catch { }
                    _toolSizes["Pen"] = Math.Clamp(_meta.StartupPenThickness > 0 ? _meta.StartupPenThickness : 1.5, 0.5, 12.0);
                    _toolSizes["Eraser"] = Math.Clamp(_meta.StartupEraserSize > 0 ? _meta.StartupEraserSize : 10.0, 4.0, 40.0);
                    _penStyleMode = string.IsNullOrEmpty(_meta.StartupPenStyle) ? "Pro" : _meta.StartupPenStyle;
                    _penThickness = _toolSizes["Pen"];
                    if (ChkExtraMargin != null) ChkExtraMargin.IsChecked = _meta.StartupExtraMargin;
                    WindowState = _meta.StartMaximized ? WindowState.Maximized : WindowState.Normal;
                    UpdatePenStudioVisuals();
                });
                firstPdf = _meta.OriginalPaths.Values.Where(File.Exists)
                    .Concat(Directory.EnumerateFiles(_libraryFolder, "*.pdf"))
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();

                if (firstPdf != null && _meta.StartupBehavior == 0)
                {
                    string f = IOPath.GetFileName(firstPdf);
                    if (!_meta.OriginalPaths.ContainsKey(f)) _meta.OriginalPaths[f] = firstPdf;
                    targetPage = _meta.LastPages.TryGetValue(f, out int lp) ? lp : 0;
                }
            });

            if (firstPdf != null && _meta.StartupBehavior == 0)
            {
                string fileName = IOPath.GetFileName(firstPdf);
                SetStartupProgress($"'{IOPath.GetFileNameWithoutExtension(fileName)}' belleğe alınıyor...", 0.60);
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);

                PdfDocument doc = null;
                var preHighRes = new Dictionary<int, BitmapSource>();
                var preThumbs = new Dictionary<int, BitmapSource>();
                var preStrokes = new Dictionary<int, byte[]>();
                int clampedPage = 0;

                await Task.Run(() =>
                {
                    doc = PdfDocument.Load(firstPdf);
                    int pageCount = doc.PageCount;
                    clampedPage = Math.Clamp(targetPage, 0, Math.Max(0, pageCount - 1));
                    var sizes = doc.PageSizes;

                    // Görünür sayfayı ve komşu sayfaları arka planda 144 DPI yüksek çözünürlükte hazırla
                    int[] warmPages = new[] { clampedPage, clampedPage + 1, clampedPage - 1 }
                        .Where(p => p >= 0 && p < pageCount).Distinct().ToArray();

                    foreach (int p in warmPages)
                    {
                        var sz = sizes[p];
                        double w = GetNormalizedPageWidth(sz), h = GetNormalizedPageHeight(sz);
                        double initScale = Math.Clamp(_meta.RenderScale > 0 ? _meta.RenderScale : 2.1, 1.5, 3.0); using var bmp = (GdiBitmap)doc.RenderCrisp(p, (int)(w * initScale), (int)(h * initScale), 144, 144, PdfRenderFlags.None);
                        preHighRes[p] = FastBitmapToSource(bmp);
                    }

                    // İlk görünür küçük resimleri (Thumbnails) arka planda hazırla
                    int tStart = Math.Max(0, clampedPage - 4);
                    int tEnd = Math.Min(pageCount - 1, clampedPage + 6);
                    for (int p = tStart; p <= tEnd; p++)
                    {
                        using var tbmp = (GdiBitmap)doc.RenderCrisp(p, 110, 152, 72, 72, PdfRenderFlags.None);
                        preThumbs[p] = FastBitmapToSource(tbmp);
                    }

                    // Çizim (.isf) dosyalarını arka planda RAM'e oku
                    string pdfKey = IOPath.GetFileNameWithoutExtension(fileName);
                    string annotDir = IOPath.Combine(_annotationsFolder, pdfKey);
                    if (Directory.Exists(annotDir))
                    {
                        foreach (var isf in Directory.EnumerateFiles(annotDir, "page_*.isf"))
                        {
                            string nm = IOPath.GetFileNameWithoutExtension(isf);
                            if (int.TryParse(nm.Replace("page_", ""), out int pIdx))
                            {
                                try { preStrokes[pIdx] = File.ReadAllBytes(isf); } catch { }
                            }
                        }
                    }
                });

                SetStartupProgress("Sayfa katmanları yerleştiriliyor...", 0.90);
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);

                var session = new PdfTabSession
                {
                    FileName = fileName,
                    FullPath = firstPdf,
                    PdfKey = IOPath.GetFileNameWithoutExtension(fileName),
                    Document = doc,
                    ZoomLevel = 1.30,
                    CurrentPage = clampedPage,
                    ExtraMarginEnabled = _meta.StartupExtraMargin || (ChkExtraMargin?.IsChecked == true),
                    PaperStyleIndex = CmbPaperStyle?.SelectedIndex ?? 0
                };

                BuildSessionSlots(session);
                foreach (var kv in preHighRes)
                    session.CacheHighRes(kv.Key, kv.Value);

                foreach (var slot in session.Slots)
                    slot.CachedStrokeBytes = preStrokes.TryGetValue(slot.PageIndex, out var b) ? b : Array.Empty<byte>();

                EnsureSessionThumbnailsInitialized(session);
                foreach (var kv in preThumbs)
                {
                    if (kv.Key >= 0 && kv.Key < session.AllThumbnails.Count)
                        session.AllThumbnails[kv.Key].ThumbnailSource = kv.Value;
                }

                _tabSessions[fileName] = session;
                if (!_openTabs.Contains(fileName)) _openTabs.Add(fileName);

                OpenPdfFile(fileName, clampedPage);
            }
            else
            {
                DashboardOverlay.Visibility = Visibility.Visible;
                RefreshDashboardTable();
            }

            SetStartupProgress("Hazır!", 1.0);
            await Task.Delay(60);

            if (StartupLoadingOverlay != null)
            {
                var fade = new System.Windows.Media.Animation.DoubleAnimation(1.0, 0.0, TimeSpan.FromMilliseconds(220));
                fade.Completed += (s, ev) => StartupLoadingOverlay.Visibility = Visibility.Collapsed;
                StartupLoadingOverlay.BeginAnimation(UIElement.OpacityProperty, fade);
            }
        }

        private void TryLoadSafeWindowIcon()
        {
            try
            {
                string icoPath = IOPath.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
                if (File.Exists(icoPath))
                    Icon = new BitmapImage(new Uri(icoPath, UriKind.Absolute));
            }
            catch { }
        }

        private string ResolvePdfPath(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;
            if (_meta.OriginalPaths.TryGetValue(fileName, out string orig) && File.Exists(orig))
                return orig;
            string libPath = IOPath.Combine(_libraryFolder, fileName);
            if (File.Exists(libPath))
                return libPath;
            return null;
        }

        // =========================================================================
        // 1. SIFIR SIKISTIRMALI HIZLI BITMAP AKTARIMI (THREAD-SAFE FROZEN BITMAP)
        // =========================================================================
        private static BitmapSource FastBitmapToSource(GdiBitmap bmp)
        {
            var rect = new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height);
            var data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            try
            {
                var src = BitmapSource.Create(
                    data.Width, data.Height, 96, 96,
                    PixelFormats.Bgra32, null,
                    data.Scan0, data.Stride * data.Height, data.Stride);
                src.Freeze();
                return src;
            }
            finally
            {
                bmp.UnlockBits(data);
            }
        }

        // =========================================================================
        // 2. 200Hz DELTA-TIME ULTRA-SMOOTH ZOOM VE KAYDIRMA MOTORU
        // =========================================================================
        private readonly Stopwatch _smoothAnimClock = Stopwatch.StartNew();
        private long _lastSmoothFrameTicks = 0;
        private double _smoothCurrentX = 0, _smoothTargetX = 0;
        private double _lastVirtualizedScrollY = -9999;

        private void EnsureRenderLoop()
        {
            if (!_renderLoopAttached)
            {
                _lastSmoothFrameTicks = _smoothAnimClock.ElapsedTicks;
                CompositionTarget.Rendering += OnRendering;
                _renderLoopAttached = true;
            }
        }

        private void StopRenderLoop()
        {
            if (_renderLoopAttached)
            {
                CompositionTarget.Rendering -= OnRendering;
                _renderLoopAttached = false;
                if (PagesStack != null && !PagesStack.UseLayoutRounding)
                    PagesStack.UseLayoutRounding = true;
            }
        }

        private void OnRendering(object sender, EventArgs e)
        {
            long nowTicks = _smoothAnimClock.ElapsedTicks;
            double dt = (double)(nowTicks - _lastSmoothFrameTicks) / Stopwatch.Frequency;
            _lastSmoothFrameTicks = nowTicks;
            dt = Math.Clamp(dt, 0.001, 0.033);

            bool active = false;

            if (_isSmoothZooming && _currentPdf != null)
            {
                // Zoom animasyonu sirasinda 1px basamaklanmayi onlemek icin LayoutRounding'i gecici kapat
                if (PagesStack != null && PagesStack.UseLayoutRounding)
                    PagesStack.UseLayoutRounding = false;

                double diffZ = _targetZoom - _zoomLevel;
                double blendZ = 1.0 - Math.Exp(-24.0 * dt);

                if (Math.Abs(diffZ) < 0.0004)
                {
                    _zoomLevel = _targetZoom;
                    _isSmoothZooming = false;
                    if (PagesStack != null) PagesStack.UseLayoutRounding = true;
                }
                else
                {
                    _zoomLevel += diffZ * blendZ;
                    active = true;
                }

                if (_activeSession != null) _activeSession.ZoomLevel = _zoomLevel;

                _suppressScrollSync = true;
                PageScale.ScaleX = _zoomLevel;
                PageScale.ScaleY = _zoomLevel;
                if (TxtZoom != null) TxtZoom.Text = $"%{(int)Math.Round(_zoomLevel * 100)}";

                double nx = Math.Max(0, (_zoomAnchorDoc.X * _zoomLevel) - _zoomAnchorViewport.X);
                double ny = Math.Max(0, (_zoomAnchorDoc.Y * _zoomLevel) - _zoomAnchorViewport.Y);
                MainScroll.ScrollToHorizontalOffset(nx);
                MainScroll.ScrollToVerticalOffset(ny);
                _smoothCurrentX = _smoothTargetX = nx;
                _smoothCurrentY = _smoothTargetY = ny;
                if (_activeSession != null) { _activeSession.ScrollOffsetX = nx; _activeSession.ScrollOffsetY = ny; }
                _suppressScrollSync = false;
            }

            if (_isSmoothScrolling && _currentPdf != null && !_isPanning && !_isSmoothZooming)
            {
                double maxY = Math.Max(0, MainScroll.ScrollableHeight);
                double maxX = Math.Max(0, MainScroll.ScrollableWidth);
                _smoothTargetY = Math.Clamp(_smoothTargetY, 0, maxY);
                _smoothTargetX = Math.Clamp(_smoothTargetX, 0, maxX);

                double diffY = _smoothTargetY - _smoothCurrentY;
                double diffX = _smoothTargetX - _smoothCurrentX;
                double blendS = 1.0 - Math.Exp(-26.0 * dt);

                _suppressScrollSync = true;

                if (Math.Abs(diffY) < 0.20 && Math.Abs(diffX) < 0.20)
                {
                    _smoothCurrentY = _smoothTargetY;
                    _smoothCurrentX = _smoothTargetX;
                    MainScroll.ScrollToVerticalOffset(_smoothCurrentY);
                    if (Math.Abs(diffX) > 0.01) MainScroll.ScrollToHorizontalOffset(_smoothCurrentX);
                    _isSmoothScrolling = false;
                }
                else
                {
                    _smoothCurrentY += diffY * blendS;
                    _smoothCurrentX += diffX * blendS;
                    MainScroll.ScrollToVerticalOffset(_smoothCurrentY);
                    if (Math.Abs(diffX) > 0.05) MainScroll.ScrollToHorizontalOffset(_smoothCurrentX);
                    active = true;
                }

                if (_activeSession != null)
                {
                    _activeSession.ScrollOffsetY = _smoothCurrentY;
                    _activeSession.ScrollOffsetX = _smoothCurrentX;
                }

                // Kaydirma esnasinda her 5ms'de bir Task.Run iptal etmek yerine yalnizca 110px gecildiginde hafif sanallastirma yap
                if (Math.Abs(_smoothCurrentY - _lastVirtualizedScrollY) > 110.0)
                {
                    _lastVirtualizedScrollY = _smoothCurrentY;
                    UpdateVisiblePagesVirtualization();
                }

                _suppressScrollSync = false;

                if (!_isSmoothScrolling)
                {
                    UpdateVisiblePagesVirtualization();
                    TriggerBackgroundPageRender();
                }
            }

            if (!active) StopRenderLoop();
        }

        private void RequestSmoothZoom(double newTarget, Point? focus = null)
        {
            double clamped = Math.Clamp(newTarget, 0.45, 3.5);
            Point newFocus = focus ?? new Point(Math.Max(100, MainScroll.ViewportWidth) * 0.5, Math.Max(100, MainScroll.ViewportHeight) * 0.5);

            // Eger kullanici ayni noktada art arda tekerlek ceviriyorsa odak noktasini sabit tut (titremeyi sifirlar)
            if (!_isSmoothZooming || (newFocus - _zoomAnchorViewport).Length > 6.0)
            {
                _zoomAnchorViewport = newFocus;
                double curX = _isSmoothScrolling ? _smoothCurrentX : MainScroll.HorizontalOffset;
                double curY = _isSmoothScrolling ? _smoothCurrentY : MainScroll.VerticalOffset;
                _zoomAnchorDoc = new Point((curX + _zoomAnchorViewport.X) / _zoomLevel, (curY + _zoomAnchorViewport.Y) / _zoomLevel);
            }

            _isSmoothScrolling = false;
            _targetZoom = clamped;
            _isSmoothZooming = true;
            EnsureRenderLoop();
            _zoomSettleTimer.Stop();
            _zoomSettleTimer.Start();
        }

        private void MainScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                RequestSmoothZoom(_targetZoom * Math.Pow(1.0014, e.Delta), e.GetPosition(MainScroll));
                e.Handled = true;
                return;
            }
            if (Keyboard.Modifiers == ModifierKeys.Shift)
            {
                if (!_isSmoothScrolling)
                {
                    _smoothCurrentX = _smoothTargetX = MainScroll.HorizontalOffset;
                    _smoothCurrentY = _smoothTargetY = MainScroll.VerticalOffset;
                }
                _smoothTargetX = Math.Clamp(_smoothTargetX - (e.Delta * 0.95), 0, Math.Max(0, MainScroll.ScrollableWidth));
                _isSmoothScrolling = true;
                EnsureRenderLoop();
                e.Handled = true;
                return;
            }

            if (!_isSmoothScrolling)
            {
                _smoothCurrentY = _smoothTargetY = MainScroll.VerticalOffset;
                _smoothCurrentX = _smoothTargetX = MainScroll.HorizontalOffset;
                _lastVirtualizedScrollY = _smoothCurrentY;
            }

            double step = e.Delta * (_meta.ScrollSpeed > 0.2 ? _meta.ScrollSpeed : 1.05) * 0.92;
            _smoothTargetY = Math.Clamp(_smoothTargetY - step, 0, Math.Max(0, MainScroll.ScrollableHeight));
            _isSmoothScrolling = true;
            EnsureRenderLoop();
            e.Handled = true;
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => RequestSmoothZoom(_targetZoom * 1.15);
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => RequestSmoothZoom(_targetZoom * 0.87);
        private void BtnZoomReset_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPdf == null || _pageSlots.Count == 0) return;
            var first = _pageSlots[0];
            double fit = Math.Clamp(Math.Min(Math.Max(400, MainScroll.ActualHeight - 16) / first.PdfHeight, Math.Max(600, MainScroll.ActualWidth - 80) / first.TotalWidth), 0.55, 1.35);
            RequestSmoothZoom(fit);
        }

        // =========================================================================
        // 3. PENCERE BOYUTU, TAM EKRAN SINIRI VE SURUKLEME
        // =========================================================================
                private void TopBar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject d)
            {
                while (d != null && d != TopTitleBarBorder)
                {
                    if (d == BtnHomeDashboard ||
                        d == BtnTopSettings ||
                        d == BtnPaperSettings ||
                        d == OpenTabsPanel ||
                        d is System.Windows.Controls.Primitives.ButtonBase ||
                        d is TextBox ||
                        d is ComboBox ||
                        d is MenuItem ||
                        (d is FrameworkElement fe && fe.Cursor == Cursors.Hand))
                    {
                        return;
                    }
                    d = d is Visual ? VisualTreeHelper.GetParent(d) : LogicalTreeHelper.GetParent(d);
                }
            }

            if (e.ClickCount == 2)
            {
                
                BtnWinMaximize_Click(sender, e);
                e.Handled = true;
                return;
            }

            if (e.ButtonState == MouseButtonState.Pressed)
            {
                e.Handled = true;
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    ReleaseCapture();
                    SendMessage(hwnd, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                }
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            UpdateWindowMaxRestoreIcon();
            ApplyMaximizedBoundsFix();
        }

        private void ApplyMaximizedBoundsFix()
        {
            if (RootWindowGrid == null) return;
            if (WindowState == WindowState.Maximized)
            {
                var border = SystemParameters.WindowResizeBorderThickness;
                RootWindowGrid.Margin = new Thickness(border.Left + 3, border.Top + 3, border.Right + 3, border.Bottom + 3);
            }
            else
            {
                RootWindowGrid.Margin = new Thickness(0);
            }
        }

        private void UpdateWindowMaxRestoreIcon()
        {
            if (WinMaxRestorePath == null) return;
            WinMaxRestorePath.Data = WindowState == WindowState.Maximized
                ? Geometry.Parse("M 1.5,3.5 L 8.5,3.5 L 8.5,10.5 L 1.5,10.5 Z M 3.5,3.5 L 3.5,1.5 L 10.5,1.5 L 10.5,8.5 L 8.5,8.5")
                : Geometry.Parse("M 2,2 L 10,2 L 10,10 L 2,10 Z");
        }

                private void BtnWinMinimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void BtnWinMaximize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            ApplyMaximizedBoundsFix();
            UpdateWindowMaxRestoreIcon();
        }
        private void BtnWinClose_Click(object sender, RoutedEventArgs e) => Close();

        // =========================================================================
        // 4. SOL SAYFA ONIZLEMELERI (ASENKRON GÖRÜNÜR ALAN RENDER MOTORU)
        // =========================================================================
        private void BtnTogglePagesDrawer_Click(object sender, RoutedEventArgs e)
        {
            _isPagesDrawerOpen = !_isPagesDrawerOpen;
            PagesDrawerColumn.Width = _isPagesDrawerOpen ? new GridLength(158) : new GridLength(0);
            if (_isPagesDrawerOpen) ScheduleThumbnailRender();
        }

        private void BtnFilterStarredPages_Click(object sender, RoutedEventArgs e)
        {
            _filterStarredPagesOnly = !_filterStarredPagesOnly;
            LeftStarFilterPath.Fill = _filterStarredPagesOnly ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FACC15")) : Brushes.Transparent;
            ApplyThumbnailFilter();
        }

        private void EnsureSessionThumbnailsInitialized(PdfTabSession session)
        {
            if (session?.Document == null || session.AllThumbnails.Count > 0) return;
            for (int i = 0; i < session.Document.PageCount; i++)
            {
                bool isStar = _meta.Bookmarks.Contains($"{session.FileName} | Sayfa {i + 1}");
                session.AllThumbnails.Add(new PageThumbnailItem
                {
                    PageIndex = i,
                    StarVisibility = isStar ? Visibility.Visible : Visibility.Collapsed
                });
            }
        }

        private void ScheduleThumbnailRender()
        {
            if (!_isPagesDrawerOpen || _activeSession?.Document == null) return;
            _thumbRenderDebounceTimer.Stop();
            _thumbRenderDebounceTimer.Start();
        }

        private void ThumbnailList_ScrollChanged(object sender, ScrollChangedEventArgs e) => ScheduleThumbnailRender();

        private ScrollViewer GetListBoxScrollViewer(DependencyObject dep)
        {
            if (dep == null) return null;
            if (dep is ScrollViewer sv) return sv;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(dep); i++)
            {
                var res = GetListBoxScrollViewer(VisualTreeHelper.GetChild(dep, i));
                if (res != null) return res;
            }
            return null;
        }

        private void TriggerBackgroundThumbnailRender()
        {
            var session = _activeSession;
            if (session?.Document == null || !_isPagesDrawerOpen || _thumbnailItems.Count == 0) return;

            int centerIdx = _currentPage;
            var sv = GetListBoxScrollViewer(ThumbnailList);
            if (sv != null && sv.ExtentHeight > 0)
            {
                // VirtualizingStackPanel item-based veya pixel-based kaydırma tespiti
                if (sv.ExtentHeight <= _thumbnailItems.Count + 5)
                    centerIdx = (int)(sv.VerticalOffset + (sv.ViewportHeight * 0.5));
                else
                    centerIdx = (int)((sv.VerticalOffset + (sv.ViewportHeight * 0.5)) / 176.0);
            }
            centerIdx = Math.Clamp(centerIdx, 0, _thumbnailItems.Count - 1);

            int start = Math.Max(0, centerIdx - 5);
            int end = Math.Min(_thumbnailItems.Count - 1, centerIdx + 6);

            var itemsToRender = new List<PageThumbnailItem>();
            for (int i = start; i <= end; i++)
            {
                if (_thumbnailItems[i].ThumbnailSource == null)
                    itemsToRender.Add(_thumbnailItems[i]);
            }

            if (itemsToRender.Count == 0) return;

            Task.Run(() =>
            {
                foreach (var item in itemsToRender)
                {
                    if (_activeSession != session || session.Document == null) break;
                    if (item.ThumbnailSource != null) continue;
                    try
                    {
                        BitmapSource src;
                        lock (session.SyncLock)
                        {
                            if (session.Document == null) break;
                            using var img = (GdiBitmap)session.Document.RenderCrisp(item.PageIndex, 96, 134, 56, 56, PdfRenderFlags.None);
                            src = FastBitmapToSource(img);
                        }
                        Dispatcher.BeginInvoke(new Action(() => item.ThumbnailSource = src), DispatcherPriority.Background);
                    }
                    catch { }
                }
            });
        }

        private void ApplyThumbnailFilter()
        {
            if (_activeSession == null) return;
            EnsureSessionThumbnailsInitialized(_activeSession);

            _suppressThumbnailSelection = true;
            _thumbnailItems.Clear();
            foreach (var item in _allThumbnailItems)
            {
                bool star = _meta.Bookmarks.Contains($"{_currentPdfFileName} | Sayfa {item.PageIndex + 1}");
                item.StarVisibility = star ? Visibility.Visible : Visibility.Collapsed;
                if (!_filterStarredPagesOnly || star) _thumbnailItems.Add(item);
            }
            TxtPagesHeader.Text = _filterStarredPagesOnly ? "★ Starred" : TrText(TranslateDynamicPattern($"Pages ({_allThumbnailItems.Count})"));
            var cur = _thumbnailItems.FirstOrDefault(t => t.PageIndex == _currentPage);
            if (cur != null)
            {
                ThumbnailList.SelectedItem = cur;
                ThumbnailList.ScrollIntoView(cur);
            }
            _suppressThumbnailSelection = false;
            ScheduleThumbnailRender();
        }

        private void ThumbnailList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_suppressThumbnailSelection && ThumbnailList.SelectedItem is PageThumbnailItem it)
                ScrollToPage(it.PageIndex);
        }

        // =========================================================================
        // 5. KITAPLIK EKRANI (ASENKRON KAPAK ONBELLEKLEME)
        // =========================================================================
        private void _InvokeDashboardToggleInternal()
        {
            Mouse.Captured?.ReleaseMouseCapture();
            Stylus.Captured?.ReleaseStylusCapture();

            if (SettingsOverlay != null && SettingsOverlay.Visibility == Visibility.Visible)
                SettingsOverlay.Visibility = Visibility.Collapsed;
            if (PenStudioPanel != null && PenStudioPanel.Visibility == Visibility.Visible)
                PenStudioPanel.Visibility = Visibility.Collapsed;
            if (PaperSettingsPopup != null && PaperSettingsPopup.Visibility == Visibility.Visible)
                PaperSettingsPopup.Visibility = Visibility.Collapsed;
            if (SelectModePopup != null && SelectModePopup.IsOpen)
                SelectModePopup.IsOpen = false;

            if (DashboardOverlay.Visibility == Visibility.Visible)
            {
                // Eger hic acik PDF yoksa Kitaplik ekranini kapatip bos siyah ekrana dusme
                if (_activeSession != null || _openTabs.Count > 0)
                    DashboardOverlay.Visibility = Visibility.Collapsed;
            }
            else
            {
                Panel.SetZIndex(DashboardOverlay, 9000);
                DashboardOverlay.Visibility = Visibility.Visible;
                RefreshDashboardTable();
            }
        }

        private void BtnHomeDashboard_Click(object sender, RoutedEventArgs e)
        {
            if (e != null) e.Handled = true;
            _InvokeDashboardToggleInternal();
        }

        private void DashboardNav_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string cat)
            {
                _dashboardCategory = cat;
                var activeBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EEF2F6"));
                BtnNavAllFiles.Background = cat == "All" ? activeBg : Brushes.Transparent;
                BtnNavStarred.Background = cat == "Starred" ? activeBg : Brushes.Transparent;
                RefreshDashboardTable();
            }
        }

        private void TxtDashboardSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TxtSearchPlaceholder != null)
                TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(TxtDashboardSearch.Text) ? Visibility.Visible : Visibility.Collapsed;
            RefreshDashboardTable();
        }

        private void BtnOpenLibraryFolder_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start(new ProcessStartInfo { FileName = _libraryFolder, UseShellExecute = true }); } catch { }
        }

        private void LoadBookCoverAsync(string fullPath, Border targetBorder)
        {
            if (_coverCache.TryGetValue(fullPath, out var cached))
            {
                targetBorder.Child = new Image { Source = cached, Stretch = Stretch.UniformToFill };
                return;
            }

            Task.Run(() =>
            {
                try
                {
                    string cacheFile = IOPath.Combine(_coversFolder, IOPath.GetFileNameWithoutExtension(fullPath) + ".jpg");
                    if (File.Exists(cacheFile))
                    {
                        var bi = new BitmapImage();
                        bi.BeginInit();
                        bi.CacheOption = BitmapCacheOption.OnLoad;
                        bi.UriSource = new Uri(cacheFile, UriKind.Absolute);
                        bi.EndInit();
                        bi.Freeze();
                        _coverCache[fullPath] = bi;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            targetBorder.Child = new Image { Source = bi, Stretch = Stretch.UniformToFill };
                        }), DispatcherPriority.Background);
                        return;
                    }

                    using var doc = PdfDocument.Load(fullPath);
                    if (doc.PageCount > 0)
                    {
                        using var img = (GdiBitmap)doc.RenderCrisp(0, 48, 64, 56, 56, PdfRenderFlags.None);
                        try { img.Save(cacheFile, System.Drawing.Imaging.ImageFormat.Jpeg); } catch { }
                        var src = FastBitmapToSource(img);
                        _coverCache[fullPath] = src;
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            targetBorder.Child = new Image { Source = src, Stretch = Stretch.UniformToFill };
                        }), DispatcherPriority.Background);
                    }
                }
                catch { }
            });
        }

        private void RefreshDashboardTable()
        {
            if (DashboardRowsPanel == null) return;
            _dashboardDirty = false;
            DashboardRowsPanel.Children.Clear();

            var files = _meta.OriginalPaths.Values.Where(File.Exists)
                .Concat(Directory.GetFiles(_libraryFolder, "*.pdf"))
                .GroupBy(p => IOPath.GetFileName(p), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .ToList();

            if (TxtTotalFilesBadge != null) TxtTotalFilesBadge.Text = files.Count.ToString();
            if (TxtStarredCountBadge != null) TxtStarredCountBadge.Text = _meta.Bookmarks.Count.ToString();

            string query = (TxtDashboardSearch?.Text ?? "").Trim().ToLowerInvariant();

            if (_dashboardCategory == "Starred")
            {
                foreach (string b in _meta.Bookmarks)
                {
                    if (!string.IsNullOrEmpty(query) && !b.ToLowerInvariant().Contains(query)) continue;

                    var row = new Border
                    {
                        Background = Brushes.White,
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                        BorderThickness = new Thickness(0, 0, 0, 1),
                        Padding = new Thickness(10, 10, 10, 10),
                        Cursor = Cursors.Hand
                    };
                    row.MouseEnter += (s, e) => row.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                    row.MouseLeave += (s, e) => row.Background = Brushes.White;

                    var g = new Grid();
                    g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });

                    var sp = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                    sp.Children.Add(new TextBlock { Text = "★  ", Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EAB308")), FontSize = 14, FontWeight = FontWeights.Bold });
                    sp.Children.Add(new TextBlock { Text = b, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F172A")), FontSize = 13, FontWeight = FontWeights.SemiBold });
                    Grid.SetColumn(sp, 0);

                    var del = new Button { Content = "Sil", Style = (Style)FindResource("LightGhostBtn"), Foreground = Brushes.Crimson, HorizontalAlignment = HorizontalAlignment.Right };
                    string capB = b;
                    del.Click += (s, ev) => { ev.Handled = true; _meta.Bookmarks.Remove(capB); SaveMetadata(); RefreshDashboardTable(); ApplyThumbnailFilter(); UpdateStarIconState(); };
                    Grid.SetColumn(del, 1);

                    g.Children.Add(sp); g.Children.Add(del);
                    row.Child = g;
                    row.MouseLeftButtonDown += (s, ev) =>
                    {
                        var parts = capB.Split('|');
                        if (parts.Length == 2 && int.TryParse(parts[1].Replace("Sayfa", "").Trim(), out int pNum))
                        {
                            DashboardOverlay.Visibility = Visibility.Collapsed;
                            OpenPdfFile(parts[0].Trim(), pNum - 1);
                        }
                    };
                    DashboardRowsPanel.Children.Add(row);
                }
                return;
            }

            foreach (string fullPath in files)
            {
                string fileName = IOPath.GetFileName(fullPath);
                if (!_meta.OriginalPaths.ContainsKey(fileName)) _meta.OriginalPaths[fileName] = fullPath;

                string tag = _meta.Tags.TryGetValue(fileName, out string tVal) ? tVal : "Ders Seç";
                string origLoc = _meta.OriginalPaths.TryGetValue(fileName, out string oVal) ? oVal : fullPath;
                string folderName = IOPath.GetFileName(IOPath.GetDirectoryName(origLoc) ?? "Ders Pdf");

                if (_dashboardCategory.StartsWith("Tag:"))
                {
                    string wanted = _dashboardCategory.Substring(4);
                    if (!string.Equals(tag, wanted, StringComparison.OrdinalIgnoreCase)) continue;
                }

                if (!string.IsNullOrEmpty(query))
                {
                    bool match = fileName.ToLowerInvariant().Contains(query) || folderName.ToLowerInvariant().Contains(query) || tag.ToLowerInvariant().Contains(query);
                    if (!match) continue;
                }

                var row = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E2E8F0")),
                    BorderThickness = new Thickness(0, 0, 0, 1),
                    Padding = new Thickness(8, 9, 8, 9),
                    Cursor = Cursors.Hand
                };
                row.MouseEnter += (s, e) => row.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                row.MouseLeave += (s, e) => row.Background = Brushes.White;

                var g = new Grid();
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2.8, GridUnitType.Star) });
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.3, GridUnitType.Star) });
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });

                var nameSp = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                var coverBorder = new Border
                {
                    Width = 24, Height = 32,
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F1F5F9")),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CBD5E1")),
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(3),
                    Margin = new Thickness(0, 0, 12, 0), ClipToBounds = true,
                    Child = new TextBlock { Text = "PDF", FontSize = 8, FontWeight = FontWeights.Bold, Foreground = Brushes.SlateGray, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
                };
                LoadBookCoverAsync(fullPath, coverBorder);

                var titleTxt = new TextBlock
                {
                    Text = fileName, FontSize = 13, FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0F172A")),
                    VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis
                };
                nameSp.Children.Add(coverBorder);
                nameSp.Children.Add(titleTxt);
                Grid.SetColumn(nameSp, 0);

                var locBtn = new Button
                {
                    Content = $"../{folderName}", Style = (Style)FindResource("LightGhostBtn"),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#475569")),
                    HorizontalAlignment = HorizontalAlignment.Left, ToolTip = origLoc
                };
                string capFull = File.Exists(origLoc) ? origLoc : fullPath;
                locBtn.Click += (s, ev) => { ev.Handled = true; try { Process.Start("explorer.exe", $"/select,\"{capFull}\""); } catch { } };
                Grid.SetColumn(locBtn, 1);

                int savedPage = _meta.LastPages.TryGetValue(fileName, out int spVal) ? (spVal + 1) : 1;
                var pageTxt = new TextBlock
                {
                    Text = $"Sayfa {savedPage}", FontSize = 12, FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB")),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(pageTxt, 2);

                var timeTxt = new TextBlock
                {
                    Text = FormatRelativeTime(File.GetLastWriteTime(fullPath)), FontSize = 12,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#64748B")),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(timeTxt, 3);

                var tagBtn = new Button
                {
                    Content = $"🏷 {tag}", Style = (Style)FindResource("LightGhostBtn"),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB")),
                    HorizontalAlignment = HorizontalAlignment.Left, ToolTip = "Ders Etiketi Değiştir"
                };
                string capFile = fileName;
                tagBtn.Click += (s, ev) =>
                {
                    ev.Handled = true;
                    int idx = Array.IndexOf(AvailableTags, tag);
                    _meta.Tags[capFile] = AvailableTags[(idx + 1) % AvailableTags.Length];
                    SaveMetadata();
                    RefreshDashboardTable();
                };
                Grid.SetColumn(tagBtn, 4);

                var delBtn = new Button
                {
                    Content = "🗑", Style = (Style)FindResource("LightGhostBtn"),
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444")),
                    HorizontalAlignment = HorizontalAlignment.Right, ToolTip = "Kitabı Listeden Kaldır"
                };
                delBtn.Click += (s, ev) =>
                {
                    ev.Handled = true;
                    if (MessageBox.Show($"'{capFile}' kitaplıktan kaldırılsın mı?", "Onay", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        CloseTabSession(capFile);
                        _meta.OriginalPaths.Remove(capFile);
                        _coverCache.TryRemove(fullPath, out _);
                        string libCopy = IOPath.Combine(_libraryFolder, capFile);
                        if (File.Exists(libCopy) && string.Equals(libCopy, fullPath, StringComparison.OrdinalIgnoreCase))
                        {
                            try { File.Delete(libCopy); } catch { }
                        }
                        SaveMetadata();
                        RebuildOpenTabsUI();
                        RefreshDashboardTable();
                    }
                };
                Grid.SetColumn(delBtn, 5);

                g.Children.Add(nameSp); g.Children.Add(locBtn); g.Children.Add(pageTxt); g.Children.Add(timeTxt); g.Children.Add(tagBtn); g.Children.Add(delBtn);
                row.Child = g;

                row.MouseLeftButtonDown += (s, ev) =>
                {
                    DashboardOverlay.Visibility = Visibility.Collapsed;
                    OpenPdfFile(capFile, _meta.LastPages.TryGetValue(capFile, out int lp) ? lp : 0);
                };

                DashboardRowsPanel.Children.Add(row);
            }
        }

        private static string FormatRelativeTime(DateTime dt)
        {
            TimeSpan diff = DateTime.Now - dt;
            if (diff.TotalMinutes < 2) return "Az önce";
            if (diff.TotalHours < 1) return $"{(int)diff.TotalMinutes} dk önce";
            if (diff.TotalHours < 24) return $"{(int)diff.TotalHours} saat önce";
            return dt.ToString("dd.MM.yyyy");
        }

        // =========================================================================
        // 6. UST SEKMELER, PEN STUDIO VE DONANIM IMLECI
        // =========================================================================
        private void BtnPaperSettings_Click(object sender, RoutedEventArgs e)
        {
            if (PaperSettingsPopup.Visibility == Visibility.Visible)
            {
                PaperSettingsPopup.Visibility = Visibility.Collapsed;
            }
            else
            {
                SyncMarginPopupControls();
                PaperSettingsPopup.Visibility = Visibility.Visible;
            }
        }

        private void SyncMarginPopupControls()
        {
            if (_meta == null) return;
            bool prev = _suppressMarginEvent;
            _suppressMarginEvent = true;
            try
            {
                int w = Math.Clamp(_meta.MarginWidthPx > 0 ? _meta.MarginWidthPx : 560, 360, 840);
                if (SliderPopupMarginWidth != null) SliderPopupMarginWidth.Value = w;
                if (TxtPopupMarginWidth != null) TxtPopupMarginWidth.Text = $"{w} px";

                string activeHex = string.IsNullOrWhiteSpace(_meta.MarginColorHex) ? "#FAFAFC" : _meta.MarginColorHex;
                if (MarginColorSwatchesGrid != null)
                {
                    var activeBorder = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6366F1"));
                    var idleBorder = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#334155"));
                    foreach (var btn in MarginColorSwatchesGrid.Children.OfType<Button>())
                    {
                        bool isMatch = string.Equals(btn.Tag as string, activeHex, StringComparison.OrdinalIgnoreCase);
                        btn.BorderBrush = isMatch ? activeBorder : idleBorder;
                        btn.BorderThickness = isMatch ? new Thickness(2.2) : new Thickness(1.2);
                    }
                }
            }
            finally
            {
                _suppressMarginEvent = prev;
            }
        }

        private void MarginColorSwatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex && !string.IsNullOrWhiteSpace(hex))
            {
                _meta.MarginColorHex = hex;
                _cachedMarginBitmap = null;
                SyncMarginPopupControls();

                if (ChkExtraMargin != null && ChkExtraMargin.IsChecked != true)
                {
                    ChkExtraMargin.IsChecked = true;
                }
                else if (_activeSession != null)
                {
                    foreach (var s in _activeSession.Slots)
                    {
                        if (s.IsRealized && s.MarginImage != null && s.ExtraMarginWidth > 0)
                            s.MarginImage.Source = GetOrCreateMarginBitmap(s.ExtraMarginWidth, (int)s.PdfHeight);
                    }
                }
                SaveMetadata();
            }
        }

        private void SliderPopupMarginWidth_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressMarginEvent || _meta == null) return;
            int newW = (int)Math.Round(e.NewValue);
            _meta.MarginWidthPx = newW;
            if (TxtPopupMarginWidth != null) TxtPopupMarginWidth.Text = $"{newW} px";

            _suppressSettingsEvents = true;
            if (SliderSetMarginWidth != null) SliderSetMarginWidth.Value = newW;
            if (TxtSetMarginWidth != null) TxtSetMarginWidth.Text = $"{newW} px";
            _suppressSettingsEvents = false;

            _cachedMarginBitmap = null;
            if (_activeSession != null && _activeSession.ExtraMarginEnabled)
                MarginOption_Changed(this, new RoutedEventArgs());
            SaveMetadata();
        }

        private void RebuildOpenTabsUI()
        {
            if (OpenTabsPanel == null) return;
            OpenTabsPanel.Children.Clear();
            foreach (string fn in _openTabs)
            {
                bool act = string.Equals(fn, _pendingActiveTab ?? _currentPdfFileName, StringComparison.OrdinalIgnoreCase);
                var tb = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(act ? "#0B0C10" : "#262936")),
                    CornerRadius = new CornerRadius(8, 8, 0, 0),
                    Padding = new Thickness(10, 4, 6, 4),
                    Margin = new Thickness(0, 4, 3, 0),
                    Cursor = Cursors.Hand
                };
                var sp = new StackPanel { Orientation = Orientation.Horizontal };
                sp.Children.Add(new TextBlock
                {
                    Text = fn,
                    Foreground = act ? Brushes.White : Brushes.LightGray,
                    FontSize = 11.5,
                    FontWeight = FontWeights.SemiBold,
                    MaxWidth = 180,
                    TextTrimming = TextTrimming.CharacterEllipsis
                });
                var cb = new Button { Content = "✕", Style = (Style)FindResource("TopGhostBtn"), Margin = new Thickness(6, 0, 0, 0), Tag = fn };
                string capFn = fn;
                cb.Click += (s, ev) =>
                {
                    ev.Handled = true;
                    CloseTabSession(capFn);
                };
                sp.Children.Add(cb);
                tb.Child = sp;
                tb.MouseLeftButtonDown += (s, ev) =>
                {
                    ev.Handled = true;
                    DashboardOverlay.Visibility = Visibility.Collapsed;
                    if (!act) OpenPdfFile(capFn, _meta.LastPages.TryGetValue(capFn, out int lp) ? lp : 0);
                };
                OpenTabsPanel.Children.Add(tb);
            }
        }
private void CloseTabSession(string fileName)
        {
            bool wasActive = string.Equals(fileName, _currentPdfFileName, StringComparison.OrdinalIgnoreCase);
            _openTabs.Remove(fileName);

            if (_tabSessions.TryGetValue(fileName, out var session))
            {
                SaveSessionStrokesSync(session);
                if (wasActive)
                {
                    PagesStack.Children.Clear();
                    _activeSession = null;
                }
                lock (session.SyncLock)
                {
                    session.Document?.Dispose();
                    session.Document = null;
                }
                session.Slots.Clear();
                session.LowResCache.Clear();
                session.HighResCache.Clear();
                session.AllThumbnails.Clear();
                _tabSessions.Remove(fileName);
            }

            RebuildOpenTabsUI();

            if (wasActive)
            {
                if (_openTabs.Count > 0)
                {
                    string nextFn = _openTabs.Last();
                    OpenPdfFile(nextFn, _meta.LastPages.TryGetValue(nextFn, out int lp) ? lp : 0);
                }
                else
                {
                    _thumbnailItems.Clear();
                    DashboardOverlay.Visibility = Visibility.Visible;
                    RefreshDashboardTable();
                }
            }
        }

        private void BtnPenStudioToggle_Click(object sender, RoutedEventArgs e) => PenStudioPanel.Visibility = PenStudioPanel.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        private void BtnClosePenStudio_Click(object sender, RoutedEventArgs e) => PenStudioPanel.Visibility = Visibility.Collapsed;
        private void EraserSubMode_Click(object sender, RoutedEventArgs e) { if (sender is Button b && b.Tag is string m) { _eraserSubMode = m; UpdatePenModeCardHighlights(); ApplyToolSettingsToAllRealizedPages(); } }
        private void PenModeCard_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string m)
            {
                _penStyleMode = m;
                UpdatePenModeCardHighlights();
                ApplyToolSettingsToAllRealizedPages();
            }
        }

        private void UpdatePenModeCardHighlights()
        {
            if (CardModeBasic == null) return;
            var activeBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252A3D"));
            var normalBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D202B"));
            var activeBd = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6366F1"));
                        foreach (var card in new[] { CardModeBasic, CardModePro, BtnSetPenBasic, BtnSetPenPro })
            {
                if (card == null) continue;
                bool isSel = string.Equals(card.Tag?.ToString(), _penStyleMode, StringComparison.OrdinalIgnoreCase);
                card.Background = isSel ? activeBg : normalBg;
                card.BorderBrush = isSel ? activeBd : Brushes.Transparent;
            }
            foreach (var eCard in new[] { CardEraserStroke, CardEraserPoint })
            {
                if (eCard == null) continue;
                bool isSel = string.Equals(eCard.Tag?.ToString(), _eraserSubMode, StringComparison.OrdinalIgnoreCase);
                eCard.Background = isSel ? activeBg : normalBg;
                eCard.BorderBrush = isSel ? activeBd : Brushes.Transparent;
            }
        }

        private void BtnToggleColorWheel_Click(object sender, RoutedEventArgs e)
        {
            if (CustomColorPickerBox == null) return;
            CustomColorPickerBox.Visibility = CustomColorPickerBox.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SatValCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingSatVal = true;
            SatValCanvas.CaptureMouse();
            UpdateSatValFromPoint(e.GetPosition(SatValCanvas));
        }

        private void SatValCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDraggingSatVal) UpdateSatValFromPoint(e.GetPosition(SatValCanvas));
        }

        private void SatValCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDraggingSatVal) { _isDraggingSatVal = false; SatValCanvas.ReleaseMouseCapture(); }
        }

        private void UpdateSatValFromPoint(Point pt)
        {
            double w = Math.Max(1, SatValCanvas.ActualWidth);
            double h = Math.Max(1, SatValCanvas.ActualHeight);
            _pickerSat = Math.Clamp(pt.X / w, 0.0, 1.0);
            _pickerVal = Math.Clamp(1.0 - (pt.Y / h), 0.0, 1.0);
            Canvas.SetLeft(SatValThumb, Math.Clamp(pt.X - 7, -4, w - 10));
            Canvas.SetTop(SatValThumb, Math.Clamp(pt.Y - 7, -4, h - 10));
            ApplyCustomHsvColor();
        }

        private void HueBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingHue = true;
            HueSpectrumBar.CaptureMouse();
            UpdateHueFromPoint(e.GetPosition(HueSpectrumBar));
        }

        private void HueBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDraggingHue) UpdateHueFromPoint(e.GetPosition(HueSpectrumBar));
        }

        private void HueBar_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDraggingHue) { _isDraggingHue = false; HueSpectrumBar.ReleaseMouseCapture(); }
        }

        private void UpdateHueFromPoint(Point pt)
        {
            double w = Math.Max(1, HueSpectrumBar.ActualWidth);
            double ratio = Math.Clamp(pt.X / w, 0.0, 1.0);
            _pickerHue = ratio * 360.0;
            Canvas.SetLeft(HueThumb, Math.Clamp(pt.X - 8, -4, w - 12));
            SatValCanvas.Background = new SolidColorBrush(HsvToColor(_pickerHue, 1.0, 1.0));
            ApplyCustomHsvColor();
        }

        private void ApplyCustomHsvColor()
        {
            Color chosen = HsvToColor(_pickerHue, _pickerSat, _pickerVal);
            _activeColor = chosen;
            if (_activeTool == "Highlighter") _highlighterColor = chosen;
            else _penColor = chosen;

            string hex = $"#{chosen.R:X2}{chosen.G:X2}{chosen.B:X2}";
            _suppressHexEvent = true;
            if (TxtHexInput != null) TxtHexInput.Text = hex;
            _suppressHexEvent = false;

            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();
        }

        private void TxtHexInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsLoaded || _suppressHexEvent || TxtHexInput == null) return;
            string raw = TxtHexInput.Text.Trim();
            if (!raw.StartsWith("#")) raw = "#" + raw;
            if (raw.Length == 7)
            {
                try
                {
                    Color parsed = (Color)ColorConverter.ConvertFromString(raw);
                    _activeColor = parsed;
                    if (_activeTool == "Highlighter") _highlighterColor = parsed;
                    else _penColor = parsed;
                    UpdatePenStudioVisuals();
                    ApplyToolSettingsToAllRealizedPages();
                }
                catch { }
            }
        }

        private static Color HsvToColor(double h, double s, double v)
        {
            double c = v * s;
            double x = c * (1 - Math.Abs((h / 60.0) % 2 - 1));
            double m = v - c;
            double r1 = 0, g1 = 0, b1 = 0;
            if (h < 60) { r1 = c; g1 = x; }
            else if (h < 120) { r1 = x; g1 = c; }
            else if (h < 180) { g1 = c; b1 = x; }
            else if (h < 240) { g1 = x; b1 = c; }
            else if (h < 300) { r1 = x; b1 = c; }
            else { r1 = c; b1 = x; }
            return Color.FromRgb((byte)Math.Round((r1 + m) * 255), (byte)Math.Round((g1 + m) * 255), (byte)Math.Round((b1 + m) * 255));
        }

        private void SliderThickness_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { if (!IsLoaded || _suppressStudioEvents) return; _penThickness = Math.Round(e.NewValue, 1); if (_toolSizes.ContainsKey(_activeTool)) _toolSizes[_activeTool] = _penThickness; UpdatePenStudioVisuals(); ApplyToolSettingsToAllRealizedPages(); }
        private void SliderOpacity_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { if (!IsLoaded || _suppressStudioEvents) return; double v = Math.Clamp(e.NewValue / 100.0, 0.15, 1.0); if (_activeTool == "Highlighter") _highlighterOpacity = v; else _penOpacity = v; UpdatePenStudioVisuals(); ApplyToolSettingsToAllRealizedPages(); }
        private void SwatchColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string hex && _activeTool != "Eraser")
            {
                _activeColor = (Color)ColorConverter.ConvertFromString(hex);
                if (_activeTool == "Highlighter") _highlighterColor = _activeColor;
                else _penColor = _activeColor;
                _suppressHexEvent = true;
                if (TxtHexInput != null) TxtHexInput.Text = hex.ToUpperInvariant();
                _suppressHexEvent = false;
                UpdatePenStudioVisuals();
                ApplyToolSettingsToAllRealizedPages();
            }
        }
                                private Color GetEffectiveInkColor()
        {
            if (_activeTool == "Highlighter")
                return GetVividHighlighterColor(_highlighterColor, _highlighterOpacity);

            return Color.FromArgb(
                (byte)Math.Clamp((int)Math.Round(_penOpacity * 255), 35, 255),
                _penColor.R, _penColor.G, _penColor.B);
        }

        private static Color GetVividHighlighterColor(Color baseCol, double opacity)
        {
            // RGB -> HSV dönüşümü ile rengin en saf, çamursuz neon tonunu (Value = 1.0) elde et
            double r = baseCol.R / 255.0;
            double g = baseCol.G / 255.0;
            double b = baseCol.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;

            double hue = 56.0; // Varsayılan: Parlak Stabilo Neon Sarı
            double sat = delta / Math.Max(0.001, max);

            if (max >= 0.28 && sat >= 0.15 && delta > 0.001)
            {
                if (max == r) hue = 60.0 * (((g - b) / delta) % 6.0);
                else if (max == g) hue = 60.0 * (((b - r) / delta) + 2.0);
                else hue = 60.0 * (((r - g) / delta) + 4.0);
                if (hue < 0) hue += 360.0;

                // Sarı tonlarının turuncu/kahverengiye kaçmasını önle, limon neon sarısına sabitle
                if (hue >= 42.0 && hue <= 64.0) hue = 56.0;
            }

            // Opaklık barı doygunluğu kontrol eder, Value (Parlaklık) her zaman %100 (1.0) kalır ki kağıt asla grileşmesin!
            double targetSat = Math.Clamp(0.35 + (Math.Clamp(opacity, 0.20, 1.0) * 0.52), 0.35, 0.87);
            return HsvToColor(hue, targetSat, 1.0);
        }

        private double GetCalibratedStrokeThickness()
        {
            if (_activeTool == "Highlighter")
                return Math.Clamp(_penThickness * 1.18, 8.0, 48.0);
            return Math.Clamp(_penThickness * 0.42, 0.18, 12.0);
        }

        private void UpdatePenStudioVisuals()
        {
            EnsurePrecisionShapeEngineHooked();
            EnsureDualSelectionEngineHooked();
            if (!IsLoaded) return;
            bool isEr = _activeTool == "Eraser", isHi = _activeTool == "Highlighter";
            Color curCol = isHi ? GetVividHighlighterColor(_highlighterColor, _highlighterOpacity) : _penColor;
            double curOp = isHi ? _highlighterOpacity : _penOpacity;

            if (TxtStudioHeader != null) TxtStudioHeader.Text = isEr ? "Eraser Studio" : (isHi ? "Highlighter Studio" : "Pen Studio");
            if (StudioHeaderDot != null) StudioHeaderDot.Fill = isEr ? Brushes.LightGray : new SolidColorBrush(curCol);
            if (AppearanceSection != null) AppearanceSection.Visibility = isEr ? Visibility.Collapsed : Visibility.Visible;
            if (PenPreviewAndModesSection != null) PenPreviewAndModesSection.Visibility = isEr ? Visibility.Collapsed : Visibility.Visible;
            if (PenModeCardsRow != null) PenModeCardsRow.Visibility = isHi ? Visibility.Collapsed : Visibility.Visible;
            if (EraserModesSection != null) EraserModesSection.Visibility = isEr ? Visibility.Visible : Visibility.Collapsed;
            if (isEr && SliderThickness != null)
            {
                _suppressStudioEvents = true;
                SliderThickness.Minimum = 4.0;
                SliderThickness.Maximum = 40.0;
                SliderThickness.TickFrequency = 1.0;
                SliderThickness.Value = Math.Clamp(_penThickness, 4.0, 40.0);
                _suppressStudioEvents = false;
            }

            if (LiveStrokePreviewPath != null && !isEr)
            {
                LiveStrokePreviewPath.Stroke = new SolidColorBrush(Color.FromArgb((byte)(curOp * 255), curCol.R, curCol.G, curCol.B));
                if (SliderThickness != null)
                {
                    _suppressStudioEvents = true;
                    if (isHi) { SliderThickness.Minimum = 6.0; SliderThickness.Maximum = 45.0; SliderThickness.TickFrequency = 1.0; }
                    else { SliderThickness.Minimum = 0.5; SliderThickness.Maximum = 12.0; SliderThickness.TickFrequency = 0.1; }
                    SliderThickness.Value = Math.Clamp(_penThickness, SliderThickness.Minimum, SliderThickness.Maximum);
                    _suppressStudioEvents = false;
                }
                LiveStrokePreviewPath.StrokeThickness = isHi
                    ? Math.Clamp(_penThickness * 0.60, 3.0, 22.0)
                    : Math.Clamp(_penThickness * 0.52, 0.35, 12.0);
            SyncDualSelectButtons();
            }

            if (CustomColorPreviewBox != null && !isEr) CustomColorPreviewBox.Background = new SolidColorBrush(curCol);
            if (TxtActiveHex != null && !isEr) TxtActiveHex.Text = $"#{curCol.R:X2}{curCol.G:X2}{curCol.B:X2}";
            if (QuickColorDot != null) QuickColorDot.Fill = isEr ? Brushes.Gray : new SolidColorBrush(curCol);
            if (PenNibFillPath != null) PenNibFillPath.Fill = new SolidColorBrush(_penColor);
            if (HighlighterNibPath != null) HighlighterNibPath.Fill = new SolidColorBrush(_highlighterColor);
            if (TxtBadgePen != null) TxtBadgePen.Text = _toolSizes["Pen"].ToString("0.#", CultureInfo.InvariantCulture);
            if (TxtBadgeHigh != null) TxtBadgeHigh.Text = ((int)_toolSizes["Highlighter"]).ToString();
            if (TxtBadgeEraser != null) TxtBadgeEraser.Text = ((int)_toolSizes["Eraser"]).ToString();
            if (TxtThicknessPt != null) TxtThicknessPt.Text = $"{_penThickness:0.#} pt";
            if (TxtOpacityPct != null) TxtOpacityPct.Text = $"{(int)(curOp * 100)}%";
            UpdatePenModeCardHighlights();
        }

        private void RebuildHardwareCursor()
        {
            if (_activeTool == "Lasso") { _hardwareCursor = Cursors.Arrow; return; }
            if (_activeTool == "Pan") { _hardwareCursor = Cursors.Hand; return; }

            string cacheKey = _activeTool == "Eraser"
                ? $"E_{(int)Math.Round(_penThickness * _zoomLevel)}"
                : (_activeTool == "Highlighter"
                    ? $"H_{(int)Math.Round(_penThickness * _zoomLevel)}_{_highlighterColor}"
                    : $"P_{_penThickness:0.#}_{_penColor}");

            if (_cursorCache.TryGetValue(cacheKey, out var cachedCursor))
            {
                _hardwareCursor = cachedCursor;
                return;
            }

            try
            {
                bool isEraser = _activeTool == "Eraser";
                int bmpSize = 64;
                int center = bmpSize / 2;

                using var bmp = new GdiBitmap(bmpSize, bmpSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var g = GdiGraphics.FromImage(bmp))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.Clear(GdiColor.Transparent);

                    if (isEraser)
                    {
                        float diameter = (float)Math.Clamp(_penThickness * _zoomLevel, 16.0, 58.0);
                        float r = diameter * 0.5f;
                        using var fill = new GdiSolidBrush(GdiColor.FromArgb(32, 100, 116, 139));
                        using var ring = new GdiPen(GdiColor.FromArgb(220, 71, 85, 105), 1.5f);
                        g.FillEllipse(fill, center - r, center - r, diameter, diameter);
                        g.DrawEllipse(ring, center - r, center - r, diameter, diameter);
                    }
                    else if (_activeTool == "Highlighter")
                    {
                        float diameter = (float)Math.Clamp(_penThickness * _zoomLevel * 0.85, 14.0, 32.0);
                        float r = diameter * 0.5f;
                        using var fill = new GdiSolidBrush(GdiColor.FromArgb(145, _highlighterColor.R, _highlighterColor.G, _highlighterColor.B));
                        using var border = new GdiPen(GdiColor.FromArgb(180, 30, 41, 59), 1.1f);
                        g.FillEllipse(fill, center - r, center - r, diameter, diameter);
                        g.DrawEllipse(border, center - r, center - r, diameter, diameter);
                    }
                    else
                    {
                        float dotDiameter = (float)Math.Clamp(GetCalibratedStrokeThickness() * _zoomLevel * 1.45, 3.0, 20.0);
                        float dotR = dotDiameter * 0.5f;
                        using var coreFill = new GdiSolidBrush(GdiColor.FromArgb(255, _penColor.R, _penColor.G, _penColor.B));
                        using var thinEdge = new GdiPen(GdiColor.FromArgb(170, 15, 23, 42), 1.1f);
                        g.FillEllipse(coreFill, center - dotR, center - dotR, dotDiameter, dotDiameter);
                        g.DrawEllipse(thinEdge, center - dotR, center - dotR, dotDiameter, dotDiameter);
                    }
                }

                IntPtr hIcon = bmp.GetHicon();
                if (GetIconInfo(hIcon, out ICONINFO info))
                {
                    info.fIcon = false;
                    info.xHotspot = center;
                    info.yHotspot = center;
                    IntPtr hCursor = CreateIconIndirect(ref info);
                    if (info.hbmColor != IntPtr.Zero) DeleteObject(info.hbmColor);
                    if (info.hbmMask != IntPtr.Zero) DeleteObject(info.hbmMask);
                    DestroyIcon(hIcon);
                    if (hCursor != IntPtr.Zero)
                    {
                        _hardwareCursor = CursorInteropHelper.Create(new SafeCursorHandle(hCursor));
                        if (_cursorCache.Count > 40) _cursorCache.Clear();
                        _cursorCache[cacheKey] = _hardwareCursor;
                        return;
                    }
                }
                DestroyIcon(hIcon);
            }
            catch { _hardwareCursor = Cursors.Cross; }
        }

        // =========================================================================
        // 7. DRAWBOARD SEKME OTURUM MOTORU + ASENKRON 2-KATMANLI LOD RENDER
        // =========================================================================
        // #region ASYNC_PDF_SWITCH
        private readonly Dictionary<string, (string Title, string Sub, double Ratio)> _tabLoadingStates = new Dictionary<string, (string, string, double)>(StringComparer.OrdinalIgnoreCase);
        private string _pendingActiveTab = null;

        private void UpdatePerTabLoaderVisibility()
        {
            if (PdfSwitchLoadingOverlay == null) return;
            string activeTab = _pendingActiveTab ?? _currentPdfFileName;
            if (!string.IsNullOrEmpty(activeTab)
                && DashboardOverlay.Visibility != Visibility.Visible
                && _tabLoadingStates.TryGetValue(activeTab, out var st))
            {
                if (TxtPdfSwitchTitle != null) TxtPdfSwitchTitle.Text = st.Title;
                if (TxtPdfSwitchSub != null) TxtPdfSwitchSub.Text = st.Sub;
                if (PdfSwitchProgressFill != null) PdfSwitchProgressFill.Width = Math.Clamp(st.Ratio, 0.15, 1.0) * 298.0;
                PdfSwitchLoadingOverlay.Opacity = 1.0;
                PdfSwitchLoadingOverlay.Visibility = Visibility.Visible;
            }
            else
            {
                PdfSwitchLoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private void SetTabSwitchProgress(string fileName, string title, string sub, double ratio)
        {
            _tabLoadingStates[fileName] = (title, sub, ratio);
            string activeTab = _pendingActiveTab ?? _currentPdfFileName;
            if (string.Equals(activeTab, fileName, StringComparison.OrdinalIgnoreCase))
                UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();
        }

        private async Task OpenPdfFileWithCompactLoaderAsync(string fileName, int targetPage, string sourcePathToCopy = null)
        {
            // Eğer bu sekme zaten arka planda yükleniyorsa sadece görünümü o sekmeye geçir
            if (_tabLoadingStates.ContainsKey(fileName))
            {
                DashboardOverlay.Visibility = Visibility.Collapsed;
                _pendingActiveTab = fileName;
                RebuildOpenTabsUI();
                PagesStack.Children.Clear();
                UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();
                return;
            }

            string shortTitle = IOPath.GetFileNameWithoutExtension(fileName);
            _pendingActiveTab = fileName;
            SetTabSwitchProgress(fileName, $"'{shortTitle}' açılıyor...", "Yeni sekmeye geçildi, sayfalar hazırlanıyor...", 0.25);

            // 1. KULLANICIYI ANINDA (0 ms) YENI SEKMEYE AT VE YÜKLEME EKRANINI GÖSTER
            DashboardOverlay.Visibility = Visibility.Collapsed;
            if (!_openTabs.Contains(fileName)) _openTabs.Add(fileName);
            RebuildOpenTabsUI();
            PagesStack.Children.Clear();
            UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();

            // Arayüzün yeni sekmeyi ve yükleme kartını ekrana çizmesi için 2 kare tanı
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            await Task.Delay(30);

            try
            {
                SaveAllRealizedStrokes();

                string fullPath = null;
                await Task.Run(() =>
                {
                    if (!string.IsNullOrEmpty(sourcePathToCopy) && File.Exists(sourcePathToCopy))
                    {
                        string dest = IOPath.Combine(_libraryFolder, fileName);
                        try { File.Copy(sourcePathToCopy, dest, true); } catch { }
                        _meta.OriginalPaths[fileName] = File.Exists(dest) ? dest : sourcePathToCopy;
                    }
                    fullPath = ResolvePdfPath(fileName);
                });

                if (string.IsNullOrEmpty(fullPath) || !File.Exists(fullPath)) return;

                SetTabSwitchProgress(fileName, $"'{shortTitle}' işleniyor...", "Yüksek çözünürlüklü sayfa katmanları yükleniyor...", 0.60);

                PdfDocument doc = null;
                var preHighRes = new Dictionary<int, BitmapSource>();
                var preThumbs = new Dictionary<int, BitmapSource>();
                var preStrokes = new Dictionary<int, byte[]>();
                int clampedPage = 0;
                double renderScale = Math.Clamp((_meta.RenderScale > 0 ? _meta.RenderScale : 2.1) * 1.35, 2.2, 3.6);

                await Task.Run(() =>
                {
                    doc = PdfDocument.Load(fullPath);
                    int pageCount = doc.PageCount;
                    clampedPage = Math.Clamp(targetPage, 0, Math.Max(0, pageCount - 1));
                    var sizes = doc.PageSizes;

                    int[] warmPages = new[] { clampedPage, clampedPage + 1, clampedPage - 1 }
                        .Where(p => p >= 0 && p < pageCount).Distinct().ToArray();

                    foreach (int p in warmPages)
                    {
                        var sz = sizes[p];
                        double w = GetNormalizedPageWidth(sz), h = GetNormalizedPageHeight(sz);
                        using var bmp = (GdiBitmap)doc.RenderCrisp(p, (int)(w * renderScale), (int)(h * renderScale), 144, 144, PdfRenderFlags.None);
                        preHighRes[p] = FastBitmapToSource(bmp);
                    }

                    int tStart = Math.Max(0, clampedPage - 4);
                    int tEnd = Math.Min(pageCount - 1, clampedPage + 6);
                    for (int p = tStart; p <= tEnd; p++)
                    {
                        using var tbmp = (GdiBitmap)doc.RenderCrisp(p, 110, 152, 72, 72, PdfRenderFlags.None);
                        preThumbs[p] = FastBitmapToSource(tbmp);
                    }

                    string pdfKey = IOPath.GetFileNameWithoutExtension(fileName);
                    string annotDir = IOPath.Combine(_annotationsFolder, pdfKey);
                    if (Directory.Exists(annotDir))
                    {
                        foreach (var isf in Directory.EnumerateFiles(annotDir, "page_*.isf"))
                        {
                            string nm = IOPath.GetFileNameWithoutExtension(isf);
                            if (int.TryParse(nm.Replace("page_", ""), out int pIdx))
                            {
                                try { preStrokes[pIdx] = File.ReadAllBytes(isf); } catch { }
                            }
                        }
                    }
                });

                SetTabSwitchProgress(fileName, $"'{shortTitle}' hazır!", "Çalışma alanına yerleştiriliyor...", 0.95);

                var session = new PdfTabSession
                {
                    FileName = fileName,
                    FullPath = fullPath,
                    PdfKey = IOPath.GetFileNameWithoutExtension(fileName),
                    Document = doc,
                    ZoomLevel = _meta.DefaultZoom > 0.5 ? _meta.DefaultZoom : 1.30,
                    CurrentPage = clampedPage,
                    ExtraMarginEnabled = _meta.StartupExtraMargin || (ChkExtraMargin?.IsChecked == true),
                    PaperStyleIndex = CmbPaperStyle?.SelectedIndex ?? 0
                };

                BuildSessionSlots(session);
                foreach (var kv in preHighRes)
                    session.CacheHighRes(kv.Key, kv.Value, maxPages: _meta.MaxCachePages > 0 ? _meta.MaxCachePages : 24);

                foreach (var slot in session.Slots)
                    slot.CachedStrokeBytes = preStrokes.TryGetValue(slot.PageIndex, out var b) ? b : Array.Empty<byte>();

                EnsureSessionThumbnailsInitialized(session);
                foreach (var kv in preThumbs)
                {
                    if (kv.Key >= 0 && kv.Key < session.AllThumbnails.Count)
                        session.AllThumbnails[kv.Key].ThumbnailSource = kv.Value;
                }

                _tabSessions[fileName] = session;
                _tabLoadingStates.Remove(fileName);

                string currentViewTab = _pendingActiveTab ?? _currentPdfFileName;
                if (string.Equals(currentViewTab, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    _pendingActiveTab = null;
                    UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();
                    OpenPdfFile(fileName, clampedPage);
                }
            }
            finally
            {
                _tabLoadingStates.Remove(fileName);
                if (string.Equals(_pendingActiveTab, fileName, StringComparison.OrdinalIgnoreCase))
                    _pendingActiveTab = null;
                UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();
            }
        }
        // #endregion ASYNC_PDF_SWITCH

        private void OpenPdfFile(string fileName, int targetPage = 0)
        {
            if (!_tabSessions.ContainsKey(fileName))
            {
                _ = OpenPdfFileWithCompactLoaderAsync(fileName, targetPage);
                return;
            }
            _pendingActiveTab = null;
            UpdatePerTabLoaderVisibility();
            ApplyHighQualityPageRenderingOptions();
            EnsurePrecisionShapeEngineHooked();string fullPath = ResolvePdfPath(fileName);
            if (fullPath == null || !File.Exists(fullPath)) return;

            StopRenderLoop();
            _isSmoothScrolling = false;
            _isSmoothZooming = false;

            // 1. Mevcut sekmenin durumunu bellekte sakla (Dispose etme, GC çağırma!)
            if (_activeSession != null)
            {
                _activeSession.ZoomLevel = _zoomLevel;
                _activeSession.ScrollOffsetX = MainScroll.HorizontalOffset;
                _activeSession.ScrollOffsetY = MainScroll.VerticalOffset;
                _activeSession.HasInitialScrollSet = true;
                _activeSession.ExtraMarginEnabled = ChkExtraMargin?.IsChecked == true;
                _activeSession.PaperStyleIndex = CmbPaperStyle?.SelectedIndex ?? 0;
            }

            if (!_openTabs.Contains(fileName)) _openTabs.Add(fileName);

            // 2. Sekme daha önce açıldıysa RAM'den 0 ms içinde geri yükle!
            bool isExistingSession = _tabSessions.TryGetValue(fileName, out var session);
            if (!isExistingSession)
            {
                session = new PdfTabSession
                {
                    FileName = fileName,
                    FullPath = fullPath,
                    PdfKey = IOPath.GetFileNameWithoutExtension(fileName),
                    Document = PdfDocument.Load(fullPath),
                    ZoomLevel = 1.30,
                    ExtraMarginEnabled = _meta.StartupExtraMargin || (ChkExtraMargin?.IsChecked == true),
                    PaperStyleIndex = CmbPaperStyle?.SelectedIndex ?? 0
                };
                session.CurrentPage = Math.Clamp(targetPage, 0, Math.Max(0, session.Document.PageCount - 1));
                BuildSessionSlots(session);
                _tabSessions[fileName] = session;
            }
            else if (targetPage > 0 && Math.Abs(session.CurrentPage - targetPage) > 0 && DashboardOverlay.Visibility == Visibility.Visible)
            {
                // Yıldızlı sorulardan belirli bir sayfaya tıklandıysa oraya git
                session.CurrentPage = Math.Clamp(targetPage, 0, Math.Max(0, session.Document.PageCount - 1));
                session.HasInitialScrollSet = false;
            }

            _activeSession = session;
            _dashboardDirty = true;

            // Sağ çözüm kağıdı durumunu sekmeye göre senkronize et
            _suppressMarginEvent = true;
            if (ChkExtraMargin != null) ChkExtraMargin.IsChecked = session.ExtraMarginEnabled;
            if (CmbPaperStyle != null) CmbPaperStyle.SelectedIndex = session.PaperStyleIndex;
            _suppressMarginEvent = false;

            // 3. Sekmenin hazır konteynerlerini PagesStack'e bağla
            _suppressScrollSync = true;
            PagesStack.Children.Clear();
            foreach (var slot in session.Slots)
            {
                PagesStack.Children.Add(slot.Container);
                if (slot.IsRealized) ApplyToolSettingsToSlot(slot);
            }

            PagesStack.PreviewMouseLeftButtonDown -= PagesStack_GapMouseDown;
            PagesStack.PreviewMouseLeftButtonDown += PagesStack_GapMouseDown;

            _zoomLevel = _targetZoom = session.ZoomLevel;
            PageScale.ScaleX = PageScale.ScaleY = _zoomLevel;
            if (TxtZoom != null) TxtZoom.Text = $"%{(int)Math.Round(_zoomLevel * 100)}";

            UpdateLayout();

            if (session.HasInitialScrollSet)
            {
                _smoothCurrentY = _smoothTargetY = session.ScrollOffsetY;
                MainScroll.ScrollToHorizontalOffset(session.ScrollOffsetX);
                MainScroll.ScrollToVerticalOffset(session.ScrollOffsetY);
                _activeSlot = session.Slots[Math.Clamp(session.CurrentPage, 0, session.Slots.Count - 1)];
            }
            else
            {
                var targetSlot = session.Slots[Math.Clamp(session.CurrentPage, 0, session.Slots.Count - 1)];
                _activeSlot = targetSlot;
                _smoothCurrentY = _smoothTargetY = targetSlot.TopOffset * _zoomLevel;
                MainScroll.ScrollToHorizontalOffset(0);
                MainScroll.ScrollToVerticalOffset(_smoothCurrentY);
                session.ScrollOffsetY = _smoothCurrentY;
                session.HasInitialScrollSet = true;
            }
            _suppressScrollSync = false;

            RebuildOpenTabsUI();
            UpdateVisiblePagesVirtualization();
            TriggerBackgroundPageRender();
            UpdateStarIconState();
            ApplyThumbnailFilter();
        }

        private void BuildSessionSlots(PdfTabSession session)
        {
            session.Slots.Clear();
            if (session?.Document == null) return;

            int extra = session.ExtraMarginEnabled ? Math.Clamp(_meta.MarginWidthPx, 360, 840) : 0;
            double top = 0;
            var pageSizes = session.Document.PageSizes;

            for (int i = 0; i < session.Document.PageCount; i++)
            {
                var sz = pageSizes[i];
                double w = GetNormalizedPageWidth(sz), h = GetNormalizedPageHeight(sz);
                var c = new Grid
                {
                    Width = w + extra,
                    Height = h,
                    Margin = new Thickness(0, 0, 0, 14),
                    ClipToBounds = true,
                    Background = new SolidColorBrush(Color.FromRgb(22, 24, 34))
                };
                session.Slots.Add(new PageSlot
                {
                    PageIndex = i,
                    PdfWidth = w,
                    PdfHeight = h,
                    ExtraMarginWidth = extra,
                    TopOffset = top,
                    Container = c
                });
                top += h + 14;
            }
        }

        private void PagesStack_GapMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject dep)
            {
                DependencyObject cur = dep;
                while (cur != null)
                {
                    if (cur is InkCanvas) return;
                    cur = VisualTreeHelper.GetParent(cur);
                }
            }

            if (_pageSlots.Count == 0 || _activeTool == "Pan" || _activeTool == "Eraser" || _activeTool == "Lasso") return;
            Point ptInStack = e.GetPosition(PagesStack);
            PageSlot nearest = _pageSlots
                .OrderBy(s => ptInStack.Y < s.TopOffset ? (s.TopOffset - ptInStack.Y) : (ptInStack.Y > s.TopOffset + s.PdfHeight ? ptInStack.Y - (s.TopOffset + s.PdfHeight) : 0))
                .FirstOrDefault();

            if (nearest != null)
            {
                if (!nearest.IsRealized) RealizePageSlot(nearest);
                if (nearest.Canvas != null)
                {
                    _activeSlot = nearest;
                    _currentPage = nearest.PageIndex;
                    _isPointerDown = true;
                    _shapeConvertedLive = false;
                    Point raw = e.GetPosition(nearest.Canvas);
                    Point clamped = new Point(Math.Clamp(raw.X, 0, nearest.TotalWidth), Math.Clamp(raw.Y, 0, nearest.PdfHeight));
                    _dragStartPoint = _lastMovePoint = clamped;
                    _rawPoints.Clear();
                    _rawPoints.Add(clamped);
                    nearest.Canvas.CaptureMouse();
                    BeginFreehandStrokeOnSlot(nearest, clamped);
                    e.Handled = true;
                }
            }
        }

        private void MainScroll_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_currentPdf == null || _pageSlots.Count == 0) return;
            if (!_isSmoothScrolling && !_isSmoothZooming)
            {
                _smoothCurrentY = _smoothTargetY = MainScroll.VerticalOffset;
                if (_activeSession != null)
                {
                    _activeSession.ScrollOffsetX = MainScroll.HorizontalOffset;
                    _activeSession.ScrollOffsetY = MainScroll.VerticalOffset;
                }
            }
            if (!_suppressScrollSync)
            {
                UpdateVisiblePagesVirtualization();
                TriggerBackgroundPageRender();
            }
        }

        private void UpdateVisiblePagesVirtualization()
        {
            if (_currentPdf == null || _pageSlots.Count == 0) return;
            double scale = Math.Max(0.25, _zoomLevel);
            double vTop = MainScroll.VerticalOffset / scale;
            double vBot = (MainScroll.VerticalOffset + Math.Max(600, MainScroll.ViewportHeight)) / scale;
            double vMid = (vTop + vBot) * 0.5;
            PageSlot center = null;
            double best = double.MaxValue;

            foreach (var s in _pageSlots)
            {
                double d = Math.Abs((s.TopOffset + s.PdfHeight * 0.5) - vMid);
                if (d < best) { best = d; center = s; }

                // Görünür alan + 1.5 sayfa ön bellekleme (Pre-fetch)
                if (s.TopOffset + s.PdfHeight >= vTop - 950 && s.TopOffset <= vBot + 950)
                {
                    if (!s.IsRealized) RealizePageSlot(s);
                }
                else if ((s.TopOffset + s.PdfHeight < vTop - 2200 || s.TopOffset > vBot + 2200) && s.IsRealized && !_isPointerDown && s != _activeSlot)
                {
                    UnrealizePageSlot(s);
                }
            }

            if (center != null && !_suppressScrollSync)
            {
                _visibleCenterPage = center.PageIndex;
                bool pageChanged = _currentPage != center.PageIndex;
                _currentPage = center.PageIndex;
                if (_activeSlot == null || !_activeSlot.IsRealized) _activeSlot = center;

                if (pageChanged)
                {
                    _suppressThumbnailSelection = true;
                    var ti = _thumbnailItems.FirstOrDefault(t => t.PageIndex == _currentPage);
                    if (ti != null && ThumbnailList.SelectedItem != ti)
                    {
                        ThumbnailList.SelectedItem = ti;
                        ThumbnailList.ScrollIntoView(ti);
                    }
                    _suppressThumbnailSelection = false;
                    _meta.LastPages[_currentPdfFileName] = _currentPage;
                    UpdateStarIconState();
                }
            }
        }

        // ANLIK SAYFA GERCEKLESTIRME (UI THREAD'DE ASLA PDF RENDER ETMEZ!)
        private void RealizePageSlot(PageSlot slot)
        {
            if (slot.IsRealized || _activeSession == null) return;

            var pdfImg = new Image
            {
                Width = slot.PdfWidth,
                Height = slot.PdfHeight,
                Stretch = Stretch.Fill,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            RenderOptions.SetBitmapScalingMode(pdfImg, BitmapScalingMode.HighQuality);

            // Önbellekte varsa 0 ms içinde doğrudan ata!
            if (_activeSession.HighResCache.TryGetValue(slot.PageIndex, out var hiBmp))
            {
                pdfImg.Source = hiBmp;
                slot.HasHighResRendered = true;
            }
            else if (_activeSession.LowResCache.TryGetValue(slot.PageIndex, out var loBmp))
            {
                pdfImg.Source = loBmp;
                slot.HasHighResRendered = false;
            }
            else
            {
                slot.HasHighResRendered = false;
            }

            var marginImg = new Image { Stretch = Stretch.None, HorizontalAlignment = HorizontalAlignment.Left };
            if (slot.ExtraMarginWidth > 0)
            {
                marginImg.Margin = new Thickness(slot.PdfWidth, 0, 0, 0);
                marginImg.Source = GetOrCreateMarginBitmap(slot.ExtraMarginWidth, (int)slot.PdfHeight);
            }

            var canvas = new InkCanvas { Width = slot.TotalWidth, Height = slot.PdfHeight, Background = Brushes.Transparent, Tag = slot };
            canvas.PreviewMouseLeftButtonDown += PageCanvas_PreviewMouseDown;
            canvas.PreviewMouseMove += PageCanvas_PreviewMouseMove;
            canvas.PreviewMouseLeftButtonUp += PageCanvas_PreviewMouseUp;
            canvas.PreviewStylusDown += PageCanvas_PreviewStylusDown;
            canvas.PreviewStylusMove += PageCanvas_PreviewStylusMove;
            canvas.PreviewStylusUp += PageCanvas_PreviewStylusUp;
            canvas.PreviewStylusButtonDown += PageCanvas_PreviewStylusButtonDown;
            canvas.PreviewStylusButtonUp += PageCanvas_PreviewStylusButtonUp;
            canvas.PreviewMouseRightButtonDown += PageCanvas_PreviewMouseRightButtonDown;
            canvas.PreviewMouseRightButtonUp += PageCanvas_PreviewMouseRightButtonUp;
            canvas.StrokeCollected += (s, e) => CommitSlotHistory(slot);
            canvas.Strokes.StrokesChanged += (s, e) => { if (!_suppressHistory) slot.HasUncommittedChanges = true; };

            slot.PdfImage = pdfImg;
            slot.MarginImage = marginImg;
            slot.Canvas = canvas;
            slot.IsRealized = true;
            slot.Container.Children.Clear();
            slot.Container.Children.Add(pdfImg);
            
            slot.Container.Children.Add(marginImg);
            var hoverCanvas = new Canvas
            {
                Width = slot.TotalWidth,
                Height = slot.PdfHeight,
                IsHitTestVisible = false,
                Background = Brushes.Transparent
            };
            slot.EraserHoverCanvas = hoverCanvas;
            slot.Container.Children.Add(canvas);
            slot.Container.Children.Add(hoverCanvas);
            
            LoadStrokesForSlot(slot);
            ApplyToolSettingsToSlot(slot);
        }

        // ARKA PLAN 2-KATMANLI (LOW-RES -> HIGH-RES) ASENKRON RENDER MOTORU
        private void TriggerBackgroundPageRender()
        {
            var session = _activeSession;
            if (session?.Document == null) return;

            var realizedSlots = session.Slots
                .Where(s => s.IsRealized && !s.HasHighResRendered)
                .OrderBy(s => Math.Abs(s.PageIndex - _visibleCenterPage))
                .ToList();

            if (realizedSlots.Count == 0) return;

            _renderCts.Cancel();
            _renderCts = new CancellationTokenSource();
            var token = _renderCts.Token;

            Task.Run(() =>
            {
                // YUKSEK COZUNURLUKLU ASENKRON RENDER: Görünür merkezden başlayarak jilet gibi net (1.55x) yüksek çözünürlüğü bas
                foreach (var slot in realizedSlots)
                {
                    if (token.IsCancellationRequested || _activeSession != session || session.Document == null) return;
                    if (!slot.IsRealized || slot.HasHighResRendered) continue;

                    try
                    {
                        if (!session.HighResCache.TryGetValue(slot.PageIndex, out BitmapSource hiSrc))
                        {
                            lock (session.SyncLock)
                            {
                                if (session.Document == null) return;
                                int hw = (int)(slot.PdfWidth * 1.55);
                                int hh = (int)(slot.PdfHeight * 1.55);
                                using var bmp = (GdiBitmap)session.Document.RenderCrisp(slot.PageIndex, hw, hh, 96, 96, PdfRenderFlags.None);
                                hiSrc = FastBitmapToSource(bmp);
                            }
                            session.CacheHighRes(slot.PageIndex, hiSrc, maxPages: _meta.MaxCachePages);
                        }

                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (slot.IsRealized && slot.PdfImage != null)
                            {
                                slot.PdfImage.Source = hiSrc;
                                slot.HasHighResRendered = true;
                            }
                        }), DispatcherPriority.Background);
                    }
                    catch { }
                }
            }, token);
        }

        private int _cachedMarginWidth = -1;
        private string _cachedMarginColor = "";

        private RenderTargetBitmap GetOrCreateMarginBitmap(int extraWidth, int height)
        {
            int style = CmbPaperStyle?.SelectedIndex ?? 0;
            string hex = string.IsNullOrWhiteSpace(_meta?.MarginColorHex) ? "#FAFAFC" : _meta.MarginColorHex;

            if (_cachedMarginBitmap != null &&
                _cachedMarginStyle == style &&
                _cachedMarginHeight == height &&
                _cachedMarginWidth == extraWidth &&
                string.Equals(_cachedMarginColor, hex, StringComparison.OrdinalIgnoreCase))
            {
                return _cachedMarginBitmap;
            }

            Color bgCol;
            try { bgCol = (Color)ColorConverter.ConvertFromString(hex); }
            catch { bgCol = Color.FromRgb(250, 250, 252); }

            bool isDarkPaper = (bgCol.R * 0.299 + bgCol.G * 0.587 + bgCol.B * 0.114) < 80;
            Color lineCol = isDarkPaper
                ? Color.FromRgb(42, 48, 70)
                : Color.FromRgb((byte)Math.Max(0, bgCol.R - 24), (byte)Math.Max(0, bgCol.G - 20), (byte)Math.Max(0, bgCol.B - 14));
            Color dotCol = isDarkPaper
                ? Color.FromRgb(68, 78, 110)
                : Color.FromRgb((byte)Math.Max(0, bgCol.R - 46), (byte)Math.Max(0, bgCol.G - 40), (byte)Math.Max(0, bgCol.B - 30));

            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                var bgBrush = new SolidColorBrush(bgCol);
                bgBrush.Freeze();
                dc.DrawRectangle(bgBrush, null, new Rect(0, 0, extraWidth, height));

                // Sol sinira ince ayirici cizgi
                var sepPen = new Pen(new SolidColorBrush(lineCol), 1.2);
                sepPen.Freeze();
                dc.DrawLine(sepPen, new Point(0.5, 0), new Point(0.5, height));

                if (style == 0)
                {
                    var p = new Pen(new SolidColorBrush(lineCol), 1.0);
                    p.Freeze();
                    for (int x = 28; x < extraWidth; x += 28) dc.DrawLine(p, new Point(x, 0), new Point(x, height));
                    for (int y = 28; y < height; y += 28) dc.DrawLine(p, new Point(0, y), new Point(extraWidth, y));
                }
                else if (style == 1)
                {
                    var dotBrush = new SolidColorBrush(dotCol);
                    dotBrush.Freeze();
                    for (int x = 24; x < extraWidth; x += 24)
                        for (int y = 24; y < height; y += 24)
                            dc.DrawEllipse(dotBrush, null, new Point(x, y), 1.3, 1.3);
                }
            }
            var rtb = new RenderTargetBitmap(extraWidth, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            _cachedMarginBitmap = rtb;
            _cachedMarginStyle = style;
            _cachedMarginHeight = height;
            _cachedMarginWidth = extraWidth;
            _cachedMarginColor = hex;
            return rtb;
        }

        private void UnrealizePageSlot(PageSlot slot)
        {
            if (!slot.IsRealized) return;
            if (slot.HasUncommittedChanges) SaveStrokesForSlot(slot, asyncDiskWrite: true);
            if (slot.PdfImage != null) slot.PdfImage.Source = null;
            if (slot.MarginImage != null) slot.MarginImage.Source = null;
            slot.Container.Children.Clear();
            slot.PdfImage = null;
            slot.EraserHoverCanvas?.Children.Clear();
            slot.EraserHoverCanvas = null;
            slot.MarginImage = null;
            slot.Canvas = null;
            slot.IsRealized = false;
            slot.HasHighResRendered = false;
        }

        private void ScrollToPage(int pageIndex)
        {
            if (_currentPdf == null || pageIndex < 0 || pageIndex >= _pageSlots.Count) return;
            _suppressScrollSync = true;
            _currentPage = pageIndex;
            _visibleCenterPage = pageIndex;
            var slot = _pageSlots[pageIndex];
            _activeSlot = slot;
            _isSmoothScrolling = false;
            _smoothCurrentY = _smoothTargetY = slot.TopOffset * _zoomLevel;
            if (_activeSession != null) _activeSession.ScrollOffsetY = _smoothCurrentY;
            MainScroll.ScrollToVerticalOffset(_smoothCurrentY);
            _suppressScrollSync = false;
            UpdateVisiblePagesVirtualization();
            TriggerBackgroundPageRender();
        }

        private void ApplyToolSettingsToAllRealizedPages()
        {
            if (_activeTool != "Eraser") ClearEraserHover();
            RebuildHardwareCursor();
            if (WorkspaceGrid != null)
                WorkspaceGrid.Cursor = _activeTool == "Pan" ? Cursors.Hand : (_activeTool == "Lasso" ? Cursors.Arrow : (_hardwareCursor ?? Cursors.Cross));

            foreach (var s in _pageSlots)
                if (s.IsRealized) ApplyToolSettingsToSlot(s);
        }

        private void ApplyToolSettingsToSlot(PageSlot slot)
        {
            if (slot?.Canvas == null) return;
            var targetCursor = _activeTool == "Pan" ? Cursors.Hand : (_activeTool == "Lasso" ? Cursors.Arrow : (_hardwareCursor ?? Cursors.Cross));
            slot.Canvas.UseCustomCursor = true;
            slot.Canvas.Cursor = targetCursor;

            InkCanvasEditingMode targetMode = _activeTool == "Lasso" ? InkCanvasEditingMode.Select : InkCanvasEditingMode.None;
            if (slot.Canvas.EditingMode != targetMode)
                slot.Canvas.EditingMode = targetMode;
        }
private string GetAnnotationPath(string pdfKey, int p)
        {
            string d = IOPath.Combine(_annotationsFolder, pdfKey ?? "Default");
            Directory.CreateDirectory(d);
            return IOPath.Combine(d, $"page_{p}.isf");
        }

        private void CommitSlotHistory(PageSlot s)
        {
            if (_suppressHistory || s?.Canvas == null) return;
            using var ms = new MemoryStream();
            s.Canvas.Strokes.Save(ms);
            byte[] bytes = ms.ToArray();
            s.CachedStrokeBytes = bytes;

            if (s.HistoryIndex < s.History.Count - 1)
                s.History.RemoveRange(s.HistoryIndex + 1, s.History.Count - s.HistoryIndex - 1);
            s.History.Add(bytes);
            if (s.History.Count > 25) s.History.RemoveAt(0);
            s.HistoryIndex = s.History.Count - 1;
            SaveStrokesForSlot(s, asyncDiskWrite: true);
        }

        private void SaveStrokesForSlot(PageSlot s, bool asyncDiskWrite = false)
        {
            if (_activeSession == null || s?.Canvas == null) return;
            s.HasUncommittedChanges = false;
            string pdfKey = _activeSession.PdfKey;
            int pageIdx = s.PageIndex;

            if (s.Canvas.Strokes.Count == 0)
            {
                s.CachedStrokeBytes = Array.Empty<byte>();
                string p = GetAnnotationPath(pdfKey, pageIdx);
                if (asyncDiskWrite)
                    Task.Run(() => { try { if (File.Exists(p)) File.Delete(p); } catch { } });
                else if (File.Exists(p))
                    File.Delete(p);
                return;
            }

            using var ms = new MemoryStream();
            s.Canvas.Strokes.Save(ms);
            byte[] data = ms.ToArray();
            s.CachedStrokeBytes = data;
            string path = GetAnnotationPath(pdfKey, pageIdx);

            if (asyncDiskWrite)
            {
                Task.Run(() =>
                {
                    try { File.WriteAllBytes(path, data); } catch { }
                });
            }
            else
            {
                try { File.WriteAllBytes(path, data); } catch { }
            }
        }

        private void SaveSessionStrokesSync(PdfTabSession session)
        {
            if (session == null) return;
            foreach (var s in session.Slots)
            {
                if (s.IsRealized && s.HasUncommittedChanges && s.Canvas != null)
                {
                    s.HasUncommittedChanges = false;
                    string path = GetAnnotationPath(session.PdfKey, s.PageIndex);
                    if (s.Canvas.Strokes.Count == 0)
                    {
                        if (File.Exists(path)) File.Delete(path);
                    }
                    else
                    {
                        using var fs = new FileStream(path, FileMode.Create);
                        s.Canvas.Strokes.Save(fs);
                    }
                }
            }
        }

        private void SaveAllRealizedStrokes()
        {
            foreach (var session in _tabSessions.Values)
                SaveSessionStrokesSync(session);
        }

        private void LoadStrokesForSlot(PageSlot s)
        {
            if (s?.Canvas == null || _activeSession == null) return;
            _suppressHistory = true;
            s.Canvas.Strokes.Clear();

            // Önce RAM önbelleğine bak (Sekme geçişinde veya kaydırmada diske gitme!)
            if (s.CachedStrokeBytes != null)
            {
                if (s.CachedStrokeBytes.Length > 0)
                {
                    try
                    {
                        using var ms = new MemoryStream(s.CachedStrokeBytes);
                        s.Canvas.Strokes.Add(new StrokeCollection(ms));
                    }
                    catch { }
                }
            }
            else
            {
                string p = GetAnnotationPath(_activeSession.PdfKey, s.PageIndex);
                if (File.Exists(p))
                {
                    try
                    {
                        byte[] diskBytes = File.ReadAllBytes(p);
                        s.CachedStrokeBytes = diskBytes;
                        using var ms = new MemoryStream(diskBytes);
                        s.Canvas.Strokes.Add(new StrokeCollection(ms));
                    }
                    catch { }
                }
                else
                {
                    s.CachedStrokeBytes = Array.Empty<byte>();
                }
            }

            _suppressHistory = false;
            s.HasUncommittedChanges = false;
            if (s.History.Count == 0)
            {
                using var ms = new MemoryStream();
                s.Canvas.Strokes.Save(ms);
                s.History.Add(ms.ToArray());
                s.HistoryIndex = 0;
            }
        }

        private PageSlot GetActiveSlot() => (_activeSlot != null && _activeSlot.IsRealized) ? _activeSlot : _pageSlots.FirstOrDefault(s => s.IsRealized);
        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            var s = GetActiveSlot();
            if (s != null && s.HistoryIndex > 0)
            {
                s.HistoryIndex--;
                _suppressHistory = true;
                s.Canvas.Strokes.Clear();
                byte[] snap = s.History[s.HistoryIndex];
                s.CachedStrokeBytes = snap;
                using var ms = new MemoryStream(snap);
                s.Canvas.Strokes.Add(new StrokeCollection(ms));
                _suppressHistory = false;
                SaveStrokesForSlot(s, asyncDiskWrite: true);
            }
        }
        private void BtnRedo_Click(object sender, RoutedEventArgs e)
        {
            var s = GetActiveSlot();
            if (s != null && s.HistoryIndex < s.History.Count - 1)
            {
                s.HistoryIndex++;
                _suppressHistory = true;
                s.Canvas.Strokes.Clear();
                byte[] snap = s.History[s.HistoryIndex];
                s.CachedStrokeBytes = snap;
                using var ms = new MemoryStream(snap);
                s.Canvas.Strokes.Add(new StrokeCollection(ms));
                _suppressHistory = false;
                SaveStrokesForSlot(s, asyncDiskWrite: true);
            }
        }
        private void BtnClearPage_Click(object sender, RoutedEventArgs e) { var s = GetActiveSlot(); if (s?.Canvas != null) { s.Canvas.Strokes.Clear(); CommitSlotHistory(s); } }
        private void BtnCopyPage_Click(object sender, RoutedEventArgs e) { var s = GetActiveSlot(); if (s?.Container != null) { var rtb = new RenderTargetBitmap((int)s.Container.ActualWidth, (int)s.Container.ActualHeight, 96, 96, PixelFormats.Pbgra32); rtb.Render(s.Container); Clipboard.SetImage(rtb); } }

        private PageSlot FindSlotAtStackPoint(Point ptInStack)
        {
            foreach (var s in _pageSlots)
            {
                if (ptInStack.Y >= s.TopOffset && ptInStack.Y <= s.TopOffset + s.PdfHeight)
                    return s;
            }
            return null;
        }

                        private void BeginFreehandStrokeOnSlot(PageSlot slot, Point posInSlot, float pressureFactor = 0.5f)
        {
            if (slot == null) return;
            if (!slot.IsRealized) RealizePageSlot(slot);
            if (slot.Canvas == null) return;

            _activeSlot = slot;
            _currentPage = slot.PageIndex;
            float clampedPressure = Math.Clamp(pressureFactor, 0.05f, 1.0f);
            _activeFreehandPoints = new StylusPointCollection
            {
                new StylusPoint(Math.Clamp(posInSlot.X, 0, slot.TotalWidth), Math.Clamp(posInSlot.Y, 0, slot.PdfHeight), clampedPressure)
            };

            bool hi = _activeTool == "Highlighter";
            bool ignorePressure = hi || _penStyleMode == "Basic";
            double calSize = GetCalibratedStrokeThickness();

            var da = new DrawingAttributes
            {
                Color = GetEffectiveInkColor(),
                Width = hi ? Math.Max(3.8, calSize * 0.28) : calSize,
                Height = hi ? Math.Max(11.0, calSize * 1.22) : calSize,
                StylusTip = hi ? StylusTip.Rectangle : StylusTip.Ellipse,
                IsHighlighter = hi,
                FitToCurve = true,
                IgnorePressure = ignorePressure
            };

            if (_penStyleMode == "Calligraphy" && !hi)
            {
                var m = new Matrix();
                m.Rotate(35.0);
                da.StylusTipTransform = m;
            }

            _activeFreehandStroke = new Stroke(_activeFreehandPoints, da);
            _suppressHistory = true;
            slot.Canvas.Strokes.Add(_activeFreehandStroke);
            _suppressHistory = false;
        }

        private readonly HashSet<Stroke> _pendingEraseStrokes = new HashSet<Stroke>();
        private PageSlot _pendingEraseSlot;
        private bool _erasedInCurrentPass = false;

        private void ClearEraserHover()
        {
            if (_pendingEraseSlot?.EraserHoverCanvas != null)
                _pendingEraseSlot.EraserHoverCanvas.Children.Clear();
            _pendingEraseStrokes.Clear();
            _pendingEraseSlot = null;
        }

        private void RefreshPendingEraseVisuals(PageSlot slot)
        {
            if (slot?.EraserHoverCanvas == null) return;
            slot.EraserHoverCanvas.Children.Clear();
            if (_pendingEraseStrokes.Count == 0) return;

            var outerHalo = new SolidColorBrush(Color.FromArgb(145, 148, 163, 184));
            var innerMask = new SolidColorBrush(Color.FromArgb(165, 255, 255, 255));

            foreach (var st in _pendingEraseStrokes)
            {
                var geo = st.GetGeometry();
                // 1. Çizginin dışında belirgin ince kontur sınırı
                slot.EraserHoverCanvas.Children.Add(new System.Windows.Shapes.Path
                {
                    Data = geo,
                    Stroke = outerHalo,
                    StrokeThickness = st.DrawingAttributes.Width + 4.5,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round,
                    StrokeLineJoin = PenLineJoin.Round
                });
                // 2. Çizgiyi Drawboard'daki gibi soluklaştıran yarı saydam katman
                slot.EraserHoverCanvas.Children.Add(new System.Windows.Shapes.Path
                {
                    Data = geo,
                    Stroke = innerMask,
                    StrokeThickness = st.DrawingAttributes.Width + 1.2,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round,
                    StrokeLineJoin = PenLineJoin.Round
                });
            }
        }

        private void RenderEraserHover(PageSlot slot, Point pt)
        {
            // Havadayken ekstra kutu çizme; Drawboard'da efekt silgi dokunup sürüklerken çalışır
        }

        private void CommitPendingEraserStrokes()
        {
            if (_pendingEraseSlot?.Canvas != null && _pendingEraseStrokes.Count > 0)
            {
                _suppressHistory = true;
                foreach (var st in _pendingEraseStrokes)
                    _pendingEraseSlot.Canvas.Strokes.Remove(st);
                _suppressHistory = false;
                _erasedInCurrentPass = true;
            }
            ClearEraserHover();
        }

        private void EraseAtSlotPoint(PageSlot slot, Point pt)
        {
            if (slot?.Canvas == null || slot.Canvas.Strokes.Count == 0) return;
            double radius = Math.Max(6.0, _penThickness * 0.65);

            if (_eraserSubMode == "Point" || !_meta.EraserAuraEnabled)
            {
                _suppressHistory = true;
                var rect = new Rect(pt.X - radius, pt.Y - radius, radius * 2, radius * 2);
                var hit = slot.Canvas.Strokes.HitTest(rect, 10);
                if (hit.Count > 0)
                {
                    slot.Canvas.Strokes.Erase(rect);
                    _erasedInCurrentPass = true;
                    slot.HasUncommittedChanges = true;
                }
                _suppressHistory = false;
                return;
            }

            if (_pendingEraseSlot != null && _pendingEraseSlot != slot)
                CommitPendingEraserStrokes();

            _pendingEraseSlot = slot;
            var hitStrokes = slot.Canvas.Strokes.HitTest(pt, radius * 1.4);
            bool addedNew = false;
            foreach (var st in hitStrokes)
            {
                if (_pendingEraseStrokes.Add(st))
                    addedNew = true;
            }

            if (addedNew)
            {
                slot.HasUncommittedChanges = true;
                RefreshPendingEraseVisuals(slot);
            }
        }
private void ActivateTempEraser()
        {
            if (_tempEraserFromButton || _activeTool == "Eraser") return;
            if (_isPointerDown && _activeFreehandStroke != null)
            {
                _holdTimer.Stop();
                FinishFreehandStrokeOnActiveSlot();
            }
            _previousToolBeforeTempEraser = _activeTool;
            _tempEraserFromButton = true;
            _activeTool = "Eraser";
            _penThickness = _toolSizes["Eraser"];
            HighlightActiveToolButtons();
            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();
        }

        private void DeactivateTempEraser()
        {
            if (!_tempEraserFromButton) return;
            CommitPendingEraserStrokes(); if (_erasedInCurrentPass && _activeSlot != null)
            {
                _erasedInCurrentPass = false;
                CommitSlotHistory(_activeSlot);
            }
            _tempEraserFromButton = false;
            _activeTool = _previousToolBeforeTempEraser;
            if (_toolSizes.TryGetValue(_activeTool, out double sz)) _penThickness = sz;
            HighlightActiveToolButtons();
            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();
        }

        private void PageCanvas_PreviewStylusButtonDown(object sender, StylusButtonEventArgs e)
        {
            if ((_meta.BarrelButtonMode == 0 && e.StylusButton.Guid == StylusPointProperties.BarrelButton.Id) ||
                (e.StylusButton.Guid == StylusPointProperties.TipButton.Id && e.StylusDevice.Inverted))
            {
                ActivateTempEraser();
            }
        }

        private void PageCanvas_PreviewStylusButtonUp(object sender, StylusButtonEventArgs e)
        {
            if (e.StylusButton.Guid == StylusPointProperties.BarrelButton.Id && _tempEraserFromButton)
                DeactivateTempEraser();
        }

        private void PageCanvas_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ActivateTempEraser();
            if (sender is InkCanvas c && c.Tag is PageSlot slot)
            {
                _activeSlot = slot;
                _isPointerDown = true;
                _erasedInCurrentPass = false;
                EraseAtSlotPoint(slot, e.GetPosition(c));
            }
            e.Handled = true;
        }

        private void PageCanvas_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            _isPointerDown = false;
            DeactivateTempEraser();
            e.Handled = true;
        }

        private void PageCanvas_PreviewStylusDown(object sender, StylusDownEventArgs e)
        {
            if (sender is not InkCanvas c || c.Tag is not PageSlot slot) return;
            if (_isQuestionCropMode) { BeginCropOnCanvas(c, e.GetPosition(c), true); e.Handled = true; return; }
            if (e.StylusDevice != null && (e.StylusDevice.Inverted || (_meta.BarrelButtonMode == 0 && e.StylusDevice.StylusButtons.Any(b => b.Guid == StylusPointProperties.BarrelButton.Id && b.StylusButtonState == StylusButtonState.Down))))
            {
                ActivateTempEraser();
            }

            if (_activeTool == "Lasso") return;

            _activeSlot = slot;
            _currentPage = slot.PageIndex;

            if (_activeTool == "Pan")
            {
                _isSmoothScrolling = false;
                _isPanning = true;
                _panStartMouse = e.GetPosition(MainScroll);
                _panStartScrollOffset = new Point(MainScroll.HorizontalOffset, MainScroll.VerticalOffset);
                _lastPanY = _panStartMouse.Y;
                _lastPanTime = DateTime.UtcNow;
                _panVelocityY = 0;
                c.CaptureStylus();
                e.Handled = true;
                return;
            }

            Point rawPos = e.GetPosition(c);
            Point pos = new Point(Math.Clamp(rawPos.X, 0, slot.TotalWidth), Math.Clamp(rawPos.Y, 0, slot.PdfHeight));

            if (_activeTool == "Eraser")
            {
                _isPointerDown = true;
                _erasedInCurrentPass = false;
                c.CaptureStylus();
                EraseAtSlotPoint(slot, pos);
                e.Handled = true;
                return;
            }

            _isPointerDown = true;
            _shapeConvertedLive = false;
            _dragStartPoint = _lastMovePoint = pos;
            _rawPoints.Clear();
            _rawPoints.Add(pos);
            _previewShapeStroke = null;

            c.CaptureStylus();
            e.Handled = true;

            if (_activeShapePreset > 0) return;

            var stylusPts = e.GetStylusPoints(c);
            float pressure = stylusPts.Count > 0 ? stylusPts[0].PressureFactor : 0.5f;
            BeginFreehandStrokeOnSlot(slot, pos, pressure);

            if ((_activeTool == "Pen" || _activeTool == "Highlighter") && ChkHoldToShape.IsChecked == true)
            {
                _holdTimer.Stop();
                _holdTimer.Start();
            }
        }

        private void PageCanvas_PreviewStylusMove(object sender, StylusEventArgs e)
        {
            if (_isQuestionCropMode) { if (_cropTargetInk != null) UpdateCropOnCanvas(e.GetPosition(_cropTargetInk)); e.Handled = true; return; }
            if (_isPanning)
            {
                Point cur = e.GetPosition(MainScroll);
                DateTime now = DateTime.UtcNow;
                double dt = Math.Max(0.004, (now - _lastPanTime).TotalSeconds);
                _panVelocityY = ((cur.Y - _lastPanY) / dt) * 0.16;
                _lastPanY = cur.Y;
                _lastPanTime = now;
                double nx = Math.Max(0, _panStartScrollOffset.X - (cur.X - _panStartMouse.X));
                double ny = Math.Clamp(_panStartScrollOffset.Y - (cur.Y - _panStartMouse.Y), 0, Math.Max(0, MainScroll.ScrollableHeight));
                _smoothCurrentY = _smoothTargetY = ny;
                MainScroll.ScrollToHorizontalOffset(nx);
                MainScroll.ScrollToVerticalOffset(ny);
                if (_activeSession != null) { _activeSession.ScrollOffsetX = nx; _activeSession.ScrollOffsetY = ny; }
                e.Handled = true;
                return;
            }

            if (e.StylusDevice != null && (e.StylusDevice.Inverted || (_meta.BarrelButtonMode == 0 && e.StylusDevice.StylusButtons.Any(b => b.Guid == StylusPointProperties.BarrelButton.Id && b.StylusButtonState == StylusButtonState.Down))))
            {
                ActivateTempEraser();
            }

            if (!_isPointerDown)
            {
                if (_activeTool == "Eraser" && sender is InkCanvas hoverCanvas && hoverCanvas.Tag is PageSlot hoverSlot)
                    RenderEraserHover(hoverSlot, e.GetPosition(hoverCanvas));
                return;
            }
            if (_activeSlot?.Canvas == null || _activeTool == "Lasso") return;

            if (_activeTool == "Eraser")
            {
                var pts = e.GetStylusPoints(_activeSlot.Canvas);
                foreach (var sp in pts)
                    EraseAtSlotPoint(_activeSlot, new Point(sp.X, sp.Y));
                if (pts.Count > 0)
                {
                    var last = pts[pts.Count - 1];
                    RenderEraserHover(_activeSlot, new Point(last.X, last.Y));
                }
                e.Handled = true;
                return;
            }

            if (_activeShapePreset > 0)
            {
                Point rawPos = e.GetPosition(_activeSlot.Canvas);
                Point pos = new Point(Math.Clamp(rawPos.X, 0, _activeSlot.TotalWidth), Math.Clamp(rawPos.Y, 0, _activeSlot.PdfHeight));
                _suppressHistory = true;
                if (_previewShapeStroke != null) _activeSlot.Canvas.Strokes.Remove(_previewShapeStroke);
                _previewShapeStroke = BuildPresetShapeStroke(_activeShapePreset, _dragStartPoint, pos);
                if (_previewShapeStroke != null) _activeSlot.Canvas.Strokes.Add(_previewShapeStroke);
                _suppressHistory = false;
                e.Handled = true;
                return;
            }

            Point ptInStack = e.GetPosition(PagesStack);
            PageSlot hoveredSlot = FindSlotAtStackPoint(ptInStack);

            if (hoveredSlot == null)
            {
                if (_activeFreehandStroke != null)
                {
                    _holdTimer.Stop();
                    FinishFreehandStrokeOnActiveSlot();
                }
                e.Handled = true;
                return;
            }

            IInputElement targetEl = hoveredSlot.Canvas ?? (IInputElement)hoveredSlot.Container;
            var stylusPts = e.GetStylusPoints(targetEl);
            if (stylusPts.Count == 0) return;

            if (_shapeConvertedLive && _previewShapeStroke != null)
            {
                if (_liveSnappedKind == "Line")
                {
                    var lastSp = stylusPts[stylusPts.Count - 1];
                    Point endPt = new Point(Math.Clamp(lastSp.X, 0, _activeSlot.TotalWidth), Math.Clamp(lastSp.Y, 0, _activeSlot.PdfHeight));
                    if (Dist(_liveSnappedLineStart, endPt) >= 8.0)
                    {
                        Point s0 = _liveSnappedLineStart;
                        Point s1 = endPt;
                        ApplyMagneticSnapToLine(_activeSlot.Canvas, ref s0, ref s1, _previewShapeStroke);
                        _previewShapeStroke.StylusPoints = CreateDenseLineStylusPoints(s0, s1);
                        SaveStrokeVerticesMetadata(_previewShapeStroke, "Line", new[] { s0, s1 });
                    }
                }
                e.Handled = true;
                return;
            }

            if (hoveredSlot != _activeSlot || _activeFreehandStroke == null)
            {
                _holdTimer.Stop();
                FinishFreehandStrokeOnActiveSlot();
                Point firstPt = new Point(Math.Clamp(stylusPts[0].X, 0, hoveredSlot.TotalWidth), Math.Clamp(stylusPts[0].Y, 0, hoveredSlot.PdfHeight));
                _rawPoints.Clear();
                _rawPoints.Add(firstPt);
                _dragStartPoint = _lastMovePoint = firstPt;
                BeginFreehandStrokeOnSlot(hoveredSlot, firstPt, stylusPts[0].PressureFactor);
                e.Handled = true;
                return;
            }

            foreach (var sp in stylusPts)
            {
                Point clampedPos = new Point(Math.Clamp(sp.X, 0, hoveredSlot.TotalWidth), Math.Clamp(sp.Y, 0, hoveredSlot.PdfHeight));
                if ((clampedPos - _lastMovePoint).Length >= _meta.StylusSmoothing)
                {
                    if (_activeTool == "Highlighter")
                    clampedPos = new Point(_lastMovePoint.X * 0.42 + clampedPos.X * 0.58, _lastMovePoint.Y * 0.42 + clampedPos.Y * 0.58);
                _activeFreehandPoints?.Add(new StylusPoint(clampedPos.X, clampedPos.Y, Math.Clamp(sp.PressureFactor, 0.05f, 1.0f)));
                    _rawPoints.Add(clampedPos);
                    _lastMovePoint = clampedPos;
                }
            }

            if ((_activeTool == "Pen" || _activeTool == "Highlighter") && ChkHoldToShape.IsChecked == true)
            {
                if ((_lastMovePoint - _dragStartPoint).Length > 3.5)
                {
                    _dragStartPoint = _lastMovePoint;
                    _holdTimer.Stop();
                    _holdTimer.Start();
                }
            }
            e.Handled = true;
        }

        private void PageCanvas_PreviewStylusUp(object sender, StylusEventArgs e)
        {
            if (_isQuestionCropMode) { if (_cropTargetInk != null) CompleteCropOnCanvas(e.GetPosition(_cropTargetInk)); e.Handled = true; return; }
            if (_isPanning)
            {
                _isPanning = false;
                Stylus.Captured?.ReleaseStylusCapture();
                if (Math.Abs(_panVelocityY) > 15)
                {
                    _smoothTargetY = Math.Clamp(_smoothCurrentY - _panVelocityY, 0, Math.Max(0, MainScroll.ScrollableHeight));
                    _isSmoothScrolling = true;
                    EnsureRenderLoop();
                }
                e.Handled = true;
                return;
            }

            if (!_isPointerDown) return;
            _isPointerDown = false;
            _holdTimer.Stop();
            Stylus.Captured?.ReleaseStylusCapture();

            if (_activeSlot?.Canvas == null) return;

            if (_activeTool == "Eraser")
            {
                CommitPendingEraserStrokes();
                if (_erasedInCurrentPass)
                {
                    _erasedInCurrentPass = false;
                    CommitSlotHistory(_activeSlot);
                }
                if (_tempEraserFromButton && (e.StylusDevice == null || !e.StylusDevice.Inverted))
                    DeactivateTempEraser();
                e.Handled = true;
                return;
            }

            if (_activeShapePreset > 0 || _shapeConvertedLive)
            {
                if (_shapeConvertedLive && _previewShapeStroke != null && _activeSlot?.Canvas != null)
                {
                    if (_liveSnappedKind == "Line")
                    {
                        var verts = GetStrokeVertices(_previewShapeStroke);
                        if (verts != null && verts.Count >= 2)
                            TryMergeCollinearLines(_activeSlot.Canvas, _previewShapeStroke, verts[0], verts[verts.Count - 1]);
                    }
                    ApplySpecialLineModifiersIfEnabled(_activeSlot.Canvas, _previewShapeStroke);
                }
                _shapeConvertedLive = false;
                _liveSnappedKind = null;
                _previewShapeStroke = null;
                _activeFreehandStroke = null;
                _activeFreehandPoints = null;
                CommitSlotHistory(_activeSlot);
                e.Handled = true;
                return;
            }

            FinishFreehandStrokeOnActiveSlot();
            e.Handled = true;
        }

        private void FinishFreehandStrokeOnActiveSlot()
        {
            if (_activeFreehandStroke != null && _activeSlot?.Canvas != null)
            {
                CommitSlotHistory(_activeSlot);
            }
            _activeFreehandStroke = null;
            _activeFreehandPoints = null;
        }

        private void PageCanvas_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return;
            if (sender is not InkCanvas c || c.Tag is not PageSlot slot) return;
            if (_isQuestionCropMode) { BeginCropOnCanvas(c, e.GetPosition(c), false); e.Handled = true; return; }
            if (_activeTool == "Lasso") return;

            _activeSlot = slot;
            _currentPage = slot.PageIndex;

            if (_activeTool == "Pan")
            {
                _isSmoothScrolling = false;
                _isPanning = true;
                _panStartMouse = e.GetPosition(MainScroll);
                _panStartScrollOffset = new Point(MainScroll.HorizontalOffset, MainScroll.VerticalOffset);
                _lastPanY = _panStartMouse.Y;
                _lastPanTime = DateTime.UtcNow;
                _panVelocityY = 0;
                c.CaptureMouse();
                e.Handled = true;
                return;
            }

            Point rawPos = e.GetPosition(c);
            Point pos = new Point(Math.Clamp(rawPos.X, 0, slot.TotalWidth), Math.Clamp(rawPos.Y, 0, slot.PdfHeight));

            if (_activeTool == "Eraser")
            {
                _isPointerDown = true;
                _erasedInCurrentPass = false;
                c.CaptureMouse();
                EraseAtSlotPoint(slot, pos);
                e.Handled = true;
                return;
            }

            _isPointerDown = true;
            _shapeConvertedLive = false;
            _dragStartPoint = _lastMovePoint = pos;
            _rawPoints.Clear();
            _rawPoints.Add(pos);
            _previewShapeStroke = null;

            c.CaptureMouse();
            e.Handled = true;

            if (_activeShapePreset > 0) return;

            BeginFreehandStrokeOnSlot(slot, pos);
            if ((_activeTool == "Pen" || _activeTool == "Highlighter") && ChkHoldToShape.IsChecked == true)
            {
                _holdTimer.Stop();
                _holdTimer.Start();
            }
        }

        private void PageCanvas_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_activeTool == "Eraser" && sender is InkCanvas hc && hc.Tag is PageSlot hs)
                RenderEraserHover(hs, e.GetPosition(hc));

            if (e.StylusDevice != null) return;
            if (_isPanning && _activeSlot?.Canvas != null)
            {
                Point cur = e.GetPosition(MainScroll);
                DateTime now = DateTime.UtcNow;
                double dt = Math.Max(0.004, (now - _lastPanTime).TotalSeconds);
                _panVelocityY = ((cur.Y - _lastPanY) / dt) * 0.16;
                _lastPanY = cur.Y;
                _lastPanTime = now;
                double nx = Math.Max(0, _panStartScrollOffset.X - (cur.X - _panStartMouse.X));
                double ny = Math.Clamp(_panStartScrollOffset.Y - (cur.Y - _panStartMouse.Y), 0, Math.Max(0, MainScroll.ScrollableHeight));
                _smoothCurrentY = _smoothTargetY = ny;
                MainScroll.ScrollToHorizontalOffset(nx);
                MainScroll.ScrollToVerticalOffset(ny);
                if (_activeSession != null) { _activeSession.ScrollOffsetX = nx; _activeSession.ScrollOffsetY = ny; }
                return;
            }

            if (!_isPointerDown)
            {
                if (_activeTool == "Eraser" && sender is InkCanvas hoverCanvas && hoverCanvas.Tag is PageSlot hoverSlot)
                    RenderEraserHover(hoverSlot, e.GetPosition(hoverCanvas));
                return;
            }
            if (_activeSlot?.Canvas == null || _activeTool == "Lasso") return;

            if (_activeTool == "Eraser")
            {
                Point ep = e.GetPosition(_activeSlot.Canvas);
                EraseAtSlotPoint(_activeSlot, ep);
                RenderEraserHover(_activeSlot, ep);
                return;
            }

            if (_activeShapePreset > 0)
            {
                Point rawPos = e.GetPosition(_activeSlot.Canvas);
                Point pos = new Point(Math.Clamp(rawPos.X, 0, _activeSlot.TotalWidth), Math.Clamp(rawPos.Y, 0, _activeSlot.PdfHeight));
                _suppressHistory = true;
                if (_previewShapeStroke != null) _activeSlot.Canvas.Strokes.Remove(_previewShapeStroke);
                _previewShapeStroke = BuildPresetShapeStroke(_activeShapePreset, _dragStartPoint, pos);
                if (_previewShapeStroke != null) _activeSlot.Canvas.Strokes.Add(_previewShapeStroke);
                _suppressHistory = false;
                return;
            }

            if (_shapeConvertedLive && _previewShapeStroke != null)
            {
                if (_liveSnappedKind == "Line")
                {
                    Point rawPos = e.GetPosition(_activeSlot.Canvas);
                    Point endPt = new Point(Math.Clamp(rawPos.X, 0, _activeSlot.TotalWidth), Math.Clamp(rawPos.Y, 0, _activeSlot.PdfHeight));
                    if (Dist(_liveSnappedLineStart, endPt) >= 8.0)
                    {
                        Point s0 = _liveSnappedLineStart;
                        Point s1 = endPt;
                        ApplyMagneticSnapToLine(_activeSlot.Canvas, ref s0, ref s1, _previewShapeStroke);
                        _previewShapeStroke.StylusPoints = CreateDenseLineStylusPoints(s0, s1);
                        SaveStrokeVerticesMetadata(_previewShapeStroke, "Line", new[] { s0, s1 });
                    }
                }
                return;
            }

            Point ptInStack = e.GetPosition(PagesStack);
            PageSlot hoveredSlot = FindSlotAtStackPoint(ptInStack);

            if (hoveredSlot == null)
            {
                if (_activeFreehandStroke != null)
                {
                    _holdTimer.Stop();
                    FinishFreehandStrokeOnActiveSlot();
                }
                return;
            }

            Point posInHovered = e.GetPosition(hoveredSlot.Canvas ?? (IInputElement)hoveredSlot.Container);
            Point clampedPos = new Point(
                Math.Clamp(posInHovered.X, 0, hoveredSlot.TotalWidth),
                Math.Clamp(posInHovered.Y, 0, hoveredSlot.PdfHeight)
            );

            if (hoveredSlot != _activeSlot || _activeFreehandStroke == null)
            {
                _holdTimer.Stop();
                FinishFreehandStrokeOnActiveSlot();
                _rawPoints.Clear();
                _rawPoints.Add(clampedPos);
                _dragStartPoint = _lastMovePoint = clampedPos;
                BeginFreehandStrokeOnSlot(hoveredSlot, clampedPos);
                return;
            }

            if ((clampedPos - _lastMovePoint).Length >= 1.2)
            {
                _activeFreehandPoints?.Add(new StylusPoint(clampedPos.X, clampedPos.Y));
                _rawPoints.Add(clampedPos);
                _lastMovePoint = clampedPos;
                if ((_activeTool == "Pen" || _activeTool == "Highlighter") && ChkHoldToShape.IsChecked == true && (_lastMovePoint - _dragStartPoint).Length > 3.5)
                {
                    _dragStartPoint = _lastMovePoint;
                    _holdTimer.Stop();
                    _holdTimer.Start();
                }
            }
        }

        private void PageCanvas_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return;
            if (_isQuestionCropMode) { if (_cropTargetInk != null) CompleteCropOnCanvas(e.GetPosition(_cropTargetInk)); e.Handled = true; return; }
            if (_isPanning)
            {
                _isPanning = false;
                Mouse.Captured?.ReleaseMouseCapture();
                if (Math.Abs(_panVelocityY) > 15)
                {
                    _smoothTargetY = Math.Clamp(_smoothCurrentY - _panVelocityY, 0, Math.Max(0, MainScroll.ScrollableHeight));
                    _isSmoothScrolling = true;
                    EnsureRenderLoop();
                }
                return;
            }

            if (!_isPointerDown) return;
            _isPointerDown = false;
            _holdTimer.Stop();
            Mouse.Captured?.ReleaseMouseCapture();

            if (_activeSlot?.Canvas == null) return;

            if (_activeTool == "Eraser")
            {
                CommitPendingEraserStrokes();
                if (_erasedInCurrentPass)
                {
                    _erasedInCurrentPass = false;
                    CommitSlotHistory(_activeSlot);
                }
                return;
            }

            if (_activeShapePreset > 0 || _shapeConvertedLive)
            {
                if (_shapeConvertedLive && _previewShapeStroke != null && _activeSlot?.Canvas != null)
                {
                    if (_liveSnappedKind == "Line")
                    {
                        var verts = GetStrokeVertices(_previewShapeStroke);
                        if (verts != null && verts.Count >= 2)
                            TryMergeCollinearLines(_activeSlot.Canvas, _previewShapeStroke, verts[0], verts[verts.Count - 1]);
                    }
                    ApplySpecialLineModifiersIfEnabled(_activeSlot.Canvas, _previewShapeStroke);
                }
                _shapeConvertedLive = false;
                _liveSnappedKind = null;
                _previewShapeStroke = null;
                _activeFreehandStroke = null;
                _activeFreehandPoints = null;
                CommitSlotHistory(_activeSlot);
                return;
            }

            FinishFreehandStrokeOnActiveSlot();
        }

        private void HoldTimer_Tick(object sender, EventArgs e)
        {
            _holdTimer.Stop();
            if (!_isPointerDown || _rawPoints.Count < 5 || _activeSlot?.Canvas == null) return;

            var cleanedPts = CleanHoldTailAndResampleFromPoints(_rawPoints);
            if (cleanedPts.Count < 4 || PolylineLength(cleanedPts) < 9.0) return;

            if (!TryBuildPrecisionGeometricStroke(
                _activeSlot.Canvas,
                cleanedPts,
                out Stroke geoStroke,
                out string shapeKind,
                out List<Point> keyVertices))
            {
                // Yazı, harf veya rastgele eğri ise çizgiye dönüştürme, olduğu gibi bırak!
                return;
            }

            _shapeConvertedLive = true;
            _liveSnappedKind = shapeKind;
            if (keyVertices != null && keyVertices.Count > 0)
                _liveSnappedLineStart = keyVertices[0];

            _suppressHistory = true;
            if (_activeFreehandStroke != null)
            {
                _activeSlot.Canvas.Strokes.Remove(_activeFreehandStroke);
                _activeFreehandStroke = null;
                _activeFreehandPoints = null;
            }
            if (_previewShapeStroke != null)
                _activeSlot.Canvas.Strokes.Remove(_previewShapeStroke);

            _previewShapeStroke = geoStroke;
            _activeSlot.Canvas.Strokes.Add(_previewShapeStroke);


            _suppressHistory = false;

            ShowSnapPulse(_activeSlot.Canvas, keyVertices);
        }

        private Stroke BuildPresetShapeStroke(int m, Point p1, Point p2)
        {
            var pts = new StylusPointCollection();
            double x1 = Math.Min(p1.X, p2.X), x2 = Math.Max(p1.X, p2.X), y1 = Math.Min(p1.Y, p2.Y), y2 = Math.Max(p1.Y, p2.Y);
            if (m == 1) { pts.Add(new StylusPoint(p1.X, p1.Y)); pts.Add(new StylusPoint(p2.X, p2.Y)); }
            else if (m == 2) { pts.Add(new StylusPoint(x1, y1)); pts.Add(new StylusPoint(x2, y1)); pts.Add(new StylusPoint(x2, y2)); pts.Add(new StylusPoint(x1, y2)); pts.Add(new StylusPoint(x1, y1)); }
            else if (m == 3) { double cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, rx = Math.Max(2, (x2 - x1) / 2), ry = Math.Max(2, (y2 - y1) / 2); for (int d = 0; d <= 360; d += 5) pts.Add(new StylusPoint(cx + rx * Math.Cos(d * Math.PI / 180), cy + ry * Math.Sin(d * Math.PI / 180))); }
            else if (m == 5) { double mx = (x1 + x2) / 2, my = (y1 + y2) / 2; pts.Add(new StylusPoint(mx, y1)); pts.Add(new StylusPoint(mx, y2)); pts.Add(new StylusPoint(mx, my)); pts.Add(new StylusPoint(x1, my)); pts.Add(new StylusPoint(x2, my)); }
            return pts.Count > 1 ? new Stroke(pts, new DrawingAttributes { Color = GetEffectiveInkColor(), Width = GetCalibratedStrokeThickness(), Height = GetCalibratedStrokeThickness() }) : null;
        }

        private void Tool_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Button b && b.Tag is string t && _activeTool != t)
            {
                Tool_Click(sender, new RoutedEventArgs());
                e.Handled = true;
            }
        }

                        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string t)
            {
                // FİNAL ÇÖZÜM: Hangi araca geçilirse geçilsin özel seçim kilit bayrağını ve modunu sıfırla!
                _isSelectToolExclusiveActive = false;
                _selectSubMode = "Lasso";

                if (_activeTool == t && _activeShapePreset == 0 && (t == "Pen" || t == "Highlighter" || t == "Eraser"))
                {
                    BtnPenStudioToggle_Click(sender, e);
                    return;
                }
                _tempEraserFromButton = false;
                _isQuestionCropMode = false;
                _activeTool = t;
                _activeShapePreset = 0;
                if (_toolSizes.TryGetValue(t, out double sz))
                {
                    _penThickness = sz;
                }
                HighlightActiveToolButtons();
                UpdatePenStudioVisuals();
                _suppressStudioEvents = true;
                if (SliderThickness != null)
                    SliderThickness.Value = Math.Clamp(_penThickness, SliderThickness.Minimum, SliderThickness.Maximum);
                if (SliderOpacity != null)
                    SliderOpacity.Value = Math.Round((_activeTool == "Highlighter" ? _highlighterOpacity : _penOpacity) * 100.0);
                _suppressStudioEvents = false;
                ApplyToolSettingsToAllRealizedPages();
            }
        }

        private void ShapePresetBtn_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string s && int.TryParse(s, out int m))
            {
                _activeShapePreset = _activeShapePreset == m ? 0 : m;
                HighlightActiveToolButtons();
                ApplyToolSettingsToAllRealizedPages();
            }
        }
private void HighlightActiveToolButtons()
        {
            if (BtnPen == null) return;
            var act = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#262B44"));
            var bd = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6366F1"));
            foreach (var b in new[] { BtnPan, BtnLasso, BtnPen, BtnHighlighter, BtnEraser, BtnShapeRect, BtnShapeCircle, BtnShapeLine, BtnShapeAxis })
                if (b != null) { b.Background = Brushes.Transparent; b.BorderBrush = Brushes.Transparent; }
            Button cur = _activeShapePreset switch { 1 => BtnShapeLine, 2 => BtnShapeRect, 3 => BtnShapeCircle, 5 => BtnShapeAxis, _ => _activeTool switch { "Pan" => BtnPan, "Lasso" => BtnLasso, "Highlighter" => BtnHighlighter, "Eraser" => BtnEraser, _ => BtnPen } };
            if (cur != null) { cur.Background = act; cur.BorderBrush = bd; } if (BtnLasso != null) BtnLasso.Tag = (cur == BtnLasso) ? "Active" : null;
        }

        private void MarginOption_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _suppressMarginEvent || _activeSession == null) return;
            int p = _currentPage;
            _activeSession.ExtraMarginEnabled = ChkExtraMargin?.IsChecked == true;
            _activeSession.PaperStyleIndex = CmbPaperStyle?.SelectedIndex ?? 0;

            foreach (var s in _activeSession.Slots)
                if (s.IsRealized) UnrealizePageSlot(s);

            BuildSessionSlots(_activeSession);
            PagesStack.Children.Clear();
            foreach (var slot in _activeSession.Slots)
                PagesStack.Children.Add(slot.Container);

            UpdateLayout();
            ScrollToPage(p);
        }

        private void BtnAddPdf_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog { Filter = "PDF (*.pdf)|*.pdf", Multiselect = true };
            if (d.ShowDialog() == true)
            {
                string last = null;
                foreach (var src in d.FileNames)
                {
                    if (!File.Exists(src)) continue;
                    string fn = IOPath.GetFileName(src);
                    string libDest = IOPath.Combine(_libraryFolder, fn);
                    try
                    {
                        if (!string.Equals(src, libDest, StringComparison.OrdinalIgnoreCase))
                            File.Copy(src, libDest, true);
                    }
                    catch { }
                    _meta.OriginalPaths[fn] = File.Exists(libDest) ? libDest : src;
                    last = fn;
                }
                _dashboardDirty = true;
                _isMetadataLoaded = true;
                SaveMetadata();
                RefreshDashboardTable();
                if (last != null) { DashboardOverlay.Visibility = Visibility.Collapsed; OpenPdfFile(last, 0); }
            }
        }

        private void BtnBookmarkPage_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentPdfFileName)) return;
            string en = $"{_currentPdfFileName} | Sayfa {_currentPage + 1}";
            if (_meta.Bookmarks.Contains(en)) _meta.Bookmarks.Remove(en); else _meta.Bookmarks.Add(en);
            _dashboardDirty = true; SaveMetadata(); UpdateStarIconState(); ApplyThumbnailFilter();
        }

        private void UpdateStarIconState()
        {
            if (StarIconPath != null && !string.IsNullOrEmpty(_currentPdfFileName))
                StarIconPath.Fill = _meta.Bookmarks.Contains($"{_currentPdfFileName} | Sayfa {_currentPage + 1}") ? Brushes.Yellow : Brushes.Transparent;
        }

                        

                private void RestoreBooksIfWiped()
    {
        try
        {
            if (_meta == null) return;
            var list = _meta.OriginalPaths;
            if (Directory.Exists(_libraryFolder))
            {
                foreach (var p in Directory.GetFiles(_libraryFolder, "*.pdf", SearchOption.AllDirectories))
                {
                    string fn = IOPath.GetFileName(p);
                    if (!list.ContainsKey(fn) || !File.Exists(list[fn]))
                        list[fn] = p;
                }
            }
        }
        catch { }
    }

    private bool _isMetadataLoaded = false;

        private void LoadMetadata()
        {
            try
            {
                if (File.Exists(_metaFilePath))
                {
                    string json = File.ReadAllText(_metaFilePath, Encoding.UTF8);
                    json = System.Text.RegularExpressions.Regex.Replace(json, "\"StartupBehavior\"\\s*:\\s*\"[^\"]+\"", "\"StartupBehavior\": 1");
                    var loaded = JsonSerializer.Deserialize<AppMetadata>(json);
                    if (loaded != null) _meta = loaded;
                }
            }
            catch { }

            _meta ??= new AppMetadata();
            _meta.OriginalPaths ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _meta.LastPages ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _meta.Tags ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _meta.Bookmarks ??= new List<string>();

            // Eklenen kitaplari koruyan yedek dosyadan (saved_pdfs.json) geri yukle
            try
            {
                string backupPathsFile = IOPath.Combine(_baseDir, "saved_pdfs.json");
                if (File.Exists(backupPathsFile))
                {
                    var savedMap = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(backupPathsFile, Encoding.UTF8));
                    if (savedMap != null)
                    {
                        foreach (var kv in savedMap)
                        {
                            if (!string.IsNullOrEmpty(kv.Value) && File.Exists(kv.Value))
                                _meta.OriginalPaths[kv.Key] = kv.Value;
                        }
                    }
                }
            }
            catch { }

            RestoreBooksIfWiped();
            _isMetadataLoaded = true;
        }

        private void SaveMetadata()
        {
            if (!_isMetadataLoaded || _meta == null) return;
            try
            {
                if (!string.IsNullOrEmpty(_currentPdfFileName))
                    _meta.LastPages[_currentPdfFileName] = _currentPage;

                RestoreBooksIfWiped();
                File.WriteAllText(_metaFilePath, JsonSerializer.Serialize(_meta), Encoding.UTF8);

                string backupPathsFile = IOPath.Combine(_baseDir, "saved_pdfs.json");
                File.WriteAllText(backupPathsFile, JsonSerializer.Serialize(_meta.OriginalPaths), Encoding.UTF8);
            }
            catch { }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape && _isQuestionCropMode) { _isQuestionCropMode = false; _cropTargetInk = null; ApplyToolSettingsToAllRealizedPages(); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.B) { ToggleSplitViewPanel(); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.E) { ExportAnnotatedPages(); e.Handled = true; return; }
            if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && e.Key == Key.S) { StartQuestionCropMode(); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Alt && e.Key == Key.D) { _isDashedLineMode = !_isDashedLineMode; UpdateGeoModifierButtonsUI(); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Alt && e.Key == Key.A) { _isVectorArrowMode = !_isVectorArrowMode; UpdateGeoModifierButtonsUI(); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Alt && e.SystemKey == Key.D1) { ApplyQuickPenPreset(0); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Alt && e.SystemKey == Key.D2) { ApplyQuickPenPreset(1); e.Handled = true; return; }
            if (Keyboard.Modifiers == ModifierKeys.Alt && e.SystemKey == Key.D3) { ApplyQuickPenPreset(2); e.Handled = true; return; }
            if (e.Key == Key.Escape && SettingsOverlay != null && SettingsOverlay.Visibility == Visibility.Visible) { BtnCloseSettings_Click(this, new RoutedEventArgs()); e.Handled = true; return; } if (e.Key == Key.F2 || (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.OemComma)) { BtnOpenSettings_Click(this, new RoutedEventArgs()); e.Handled = true; return; } if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Z) { BtnUndo_Click(this, new RoutedEventArgs()); e.Handled = true; }
            else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Y) { BtnRedo_Click(this, new RoutedEventArgs()); e.Handled = true; }
            else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Tab && _openTabs.Count > 1)
            {
                int idx = _openTabs.IndexOf(_currentPdfFileName);
                string next = _openTabs[(idx + 1) % _openTabs.Count];
                OpenPdfFile(next, _meta.LastPages.TryGetValue(next, out int lp) ? lp : 0);
                e.Handled = true;
            }
        }

                // #region SETTINGS_ENGINE
        private bool _suppressSettingsEvents = false;

        private void ApplyLoadedSettingsToRuntime()
        {
            _suppressSettingsEvents = true;
            try
            {
                if (_meta.HoldToLineMs < 500) _meta.HoldToLineMs = 550;
                if (_meta.StylusSmoothing < 0.2) _meta.StylusSmoothing = 0.8;
                if (_meta.RenderScale < 1.2) _meta.RenderScale = 2.1;
                if (_meta.MaxCachePages < 8) _meta.MaxCachePages = 24;
                if (_meta.ScrollSpeed < 0.3) _meta.ScrollSpeed = 1.05;
                if (_meta.DefaultZoom < 0.5) _meta.DefaultZoom = 1.30;
                if (_meta.MarginWidthPx < 300) _meta.MarginWidthPx = 560;

                _holdTimer.Interval = TimeSpan.FromMilliseconds(_meta.HoldToLineMs);
                _isPagesDrawerOpen = _meta.StartWithDrawerOpen;
                if (PagesDrawerColumn != null)
                    PagesDrawerColumn.Width = _isPagesDrawerOpen ? new GridLength(158) : new GridLength(0);

                if (SliderSetHoldMs != null) SliderSetHoldMs.Value = _meta.HoldToLineMs;
                if (TxtSetHoldMs != null) TxtSetHoldMs.Text = $"{_meta.HoldToLineMs} ms";

                if (SliderSetSmoothing != null) SliderSetSmoothing.Value = _meta.StylusSmoothing;
                if (TxtSetSmoothing != null) TxtSetSmoothing.Text = $"{_meta.StylusSmoothing:0.0} px";

                if (CmbSetBarrelMode != null) CmbSetBarrelMode.SelectedIndex = Math.Clamp(_meta.BarrelButtonMode, 0, 1);
                if (ChkSetEraserAura != null) ChkSetEraserAura.IsChecked = _meta.EraserAuraEnabled;

                if (SliderSetCachePages != null) SliderSetCachePages.Value = _meta.MaxCachePages;
                if (TxtSetCachePages != null) TxtSetCachePages.Text = $"{_meta.MaxCachePages} Sayfa";

                if (SliderSetScrollSpeed != null) SliderSetScrollSpeed.Value = _meta.ScrollSpeed;
                if (TxtSetScrollSpeed != null) TxtSetScrollSpeed.Text = $"{_meta.ScrollSpeed:0.00}x";

                if (SliderSetDefZoom != null) SliderSetDefZoom.Value = Math.Round(_meta.DefaultZoom * 100);
                if (TxtSetDefZoom != null) TxtSetDefZoom.Text = $"%{(int)Math.Round(_meta.DefaultZoom * 100)}";

                if (SliderSetMarginWidth != null) SliderSetMarginWidth.Value = _meta.MarginWidthPx;
                if (TxtSetMarginWidth != null) TxtSetMarginWidth.Text = $"{_meta.MarginWidthPx} px";

                if (ChkSetStartDrawer != null) ChkSetStartDrawer.IsChecked = _meta.StartWithDrawerOpen;

                if (string.IsNullOrEmpty(_meta.StartupPenColorHex)) _meta.StartupPenColorHex = "#0F172A";
                if (_meta.StartupPenThickness < 0.5) _meta.StartupPenThickness = 1.5;
                if (_meta.StartupEraserSize < 4.0) _meta.StartupEraserSize = 10.0;

                if (string.IsNullOrEmpty(_meta.StartupPenStyle)) _meta.StartupPenStyle = "Pro";
                if (CmbSetStartupPenStyle != null)
                    CmbSetStartupPenStyle.SelectedIndex = string.Equals(_meta.StartupPenStyle, "Basic", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                UpdatePenModeCardHighlights();
                if (CmbSetStartupAction != null) CmbSetStartupAction.SelectedIndex = Math.Clamp(_meta.StartupBehavior, 0, 1);
                if (CmbSetStartupColor != null)
                {
                    int cIdx = _meta.StartupPenColorHex.ToUpperInvariant() switch
                    {
                        "#1E3A8A" => 1,
                        "#DC2626" => 2,
                        "#2563EB" => 3,
                        _ => 0
                    };
                    CmbSetStartupColor.SelectedIndex = cIdx;
                }
                if (SliderSetStartPenSize != null)
                {
                    SliderSetStartPenSize.Minimum = 0.5;
                    SliderSetStartPenSize.Maximum = 12.0;
                    SliderSetStartPenSize.TickFrequency = 0.1;
                    SliderSetStartPenSize.IsSnapToTickEnabled = true;
                    SliderSetStartPenSize.Value = Math.Clamp(_meta.StartupPenThickness, 0.5, 12.0);
                }
                if (SliderSetStartEraserSize != null)
                {
                    SliderSetStartEraserSize.Minimum = 4.0;
                    SliderSetStartEraserSize.Maximum = 40.0;
                    SliderSetStartEraserSize.TickFrequency = 1.0;
                    SliderSetStartEraserSize.IsSnapToTickEnabled = true;
                }
                if (TxtSetStartPenSize != null) TxtSetStartPenSize.Text = $"{_meta.StartupPenThickness:0.0} pt";

                if (SliderSetStartEraserSize != null) SliderSetStartEraserSize.Value = _meta.StartupEraserSize;
                if (TxtSetStartEraserSize != null) TxtSetStartEraserSize.Text = $"{(int)_meta.StartupEraserSize} pt";

                if (ChkSetStartExtraMargin != null) ChkSetStartExtraMargin.IsChecked = _meta.StartupExtraMargin;
                if (ChkSetStartMaximized != null) ChkSetStartMaximized.IsChecked = _meta.StartMaximized;

                UpdateQualityCardHighlights();
                UpdatePageFilterHighlights();
                ApplyPageFilterEffect();
            }
            finally
            {
                _suppressSettingsEvents = false;
            }
        }

        private void BtnOpenSettings_Click(object sender, RoutedEventArgs e)
        {
            if (SettingsOverlay == null) return;
            EnsureLanguageSelectorInSettings();
            NormalizeSettingsOverlayLayout();
            ApplyApplicationLanguage(_currentLang, saveToMeta: false);
            ApplyLoadedSettingsToRuntime();
            RefreshStorageStatistics();
            SettingsOverlay.Visibility = SettingsOverlay.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        }

        private void BtnCloseSettings_Click(object sender, RoutedEventArgs e)
        {
            if (SettingsOverlay != null) SettingsOverlay.Visibility = Visibility.Collapsed;
            SaveMetadata();
        }

        private void SettingsBackdrop_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (SettingsOverlay != null) SettingsOverlay.Visibility = Visibility.Collapsed;
            SaveMetadata();
        }

        private void SettingsCard_MouseDown(object sender, MouseButtonEventArgs e) => e.Handled = true;

        private void SettingsNav_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn) return;
            foreach (var b in new[] { BtnSetTabPen, BtnSetTabPerf, BtnSetTabWork, BtnSetTabStorage, BtnSetTabStartup })
                if (b != null) b.Tag = b == btn ? "Active" : null;

            SetPagePen.Visibility = btn == BtnSetTabPen ? Visibility.Visible : Visibility.Collapsed;
            SetPagePerf.Visibility = btn == BtnSetTabPerf ? Visibility.Visible : Visibility.Collapsed;
            SetPageWork.Visibility = btn == BtnSetTabWork ? Visibility.Visible : Visibility.Collapsed;
            SetPageStorage.Visibility = btn == BtnSetTabStorage ? Visibility.Visible : Visibility.Collapsed;
            if (SetPageStartup != null) SetPageStartup.Visibility = btn == BtnSetTabStartup ? Visibility.Visible : Visibility.Collapsed;

            if (btn == BtnSetTabPen)
            {
                TxtSettingsPageTitle.Text = "Kalem ve Silgi Motoru";
                TxtSettingsPageSub.Text = "Veikk grafik tablet tepkimesi, düz çizgi süresi ve silgi ön izleme ayarları";
            }
            else if (btn == BtnSetTabPerf)
            {
                TxtSettingsPageTitle.Text = "Görüntü ve Performans";
                TxtSettingsPageSub.Text = "PDF supersampling çözünürlüğü, RAM doku önbelleği ve kaydırma hızı";
            }
            else if (btn == BtnSetTabWork)
            {
                TxtSettingsPageTitle.Text = "Çalışma Alanı ve Kağıt";
                TxtSettingsPageSub.Text = "Göz koruma filtreleri, sağ çözüm kağıdı genişliği ve panel düzeni";
            }
            else if (btn == BtnSetTabStorage)
            {
                TxtSettingsPageTitle.Text = "Depolama ve Yedekleme";
                TxtSettingsPageSub.Text = "Çözüm verileri istatistikleri, tek tıkla yedekleme ve önbellek yönetimi";
                RefreshStorageStatistics();
            }
            else if (btn == BtnSetTabStartup)
            {
                TxtSettingsPageTitle.Text = "Başlangıç Ayarları";
                TxtSettingsPageSub.Text = "Açılış ekranı davranışı, varsayılan kalem/silgi değerleri ve pencere düzeni";
            }
        }

                private void BtnSetPenStyle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string m)
            {
                _penStyleMode = m;
                _meta.StartupPenStyle = m;
                UpdatePenModeCardHighlights();
                ApplyLoadedSettingsToRuntime();
                ApplyToolSettingsToAllRealizedPages();
                SaveMetadata();
            }
        }

        private void CmbSetStartupPenStyle_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || CmbSetStartupPenStyle?.SelectedItem is not ComboBoxItem item || item.Tag is not string style) return;
            _meta.StartupPenStyle = style;
            _penStyleMode = style;
            UpdatePenModeCardHighlights();
            ApplyToolSettingsToAllRealizedPages();
            SaveMetadata();
        }

        private void CmbSetStartupAction_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || CmbSetStartupAction == null) return;
            _meta.StartupBehavior = CmbSetStartupAction.SelectedIndex;
            SaveMetadata();
        }

        private void CmbSetStartupColor_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || CmbSetStartupColor?.SelectedItem is not ComboBoxItem item || item.Tag is not string hex) return;
            _meta.StartupPenColorHex = hex;
            _penColor = _activeColor = (Color)ColorConverter.ConvertFromString(hex);
            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();
            SaveMetadata();
        }

        private void SliderSetStartPenSize_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.StartupPenThickness = Math.Round(e.NewValue, 1);
            _toolSizes["Pen"] = _meta.StartupPenThickness;
            if (_activeTool == "Pen") _penThickness = _meta.StartupPenThickness;
            if (TxtSetStartPenSize != null) TxtSetStartPenSize.Text = $"{_meta.StartupPenThickness:0.0} pt";
            UpdatePenStudioVisuals();
            SaveMetadata();
        }

        private void SliderSetStartEraserSize_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.StartupEraserSize = Math.Round(e.NewValue);
            _toolSizes["Eraser"] = _meta.StartupEraserSize;
            if (_activeTool == "Eraser") _penThickness = _meta.StartupEraserSize;
            if (TxtSetStartEraserSize != null) TxtSetStartEraserSize.Text = $"{(int)_meta.StartupEraserSize} pt";
            UpdatePenStudioVisuals();
            SaveMetadata();
        }

        private void ChkSetStartExtraMargin_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || ChkSetStartExtraMargin == null) return;
            _meta.StartupExtraMargin = ChkSetStartExtraMargin.IsChecked == true;
            SaveMetadata();
        }

        private void ChkSetStartMaximized_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || ChkSetStartMaximized == null) return;
            _meta.StartMaximized = ChkSetStartMaximized.IsChecked == true;
            SaveMetadata();
        }
private void SliderSetHoldMs_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.HoldToLineMs = (int)Math.Round(e.NewValue);
            _holdTimer.Interval = TimeSpan.FromMilliseconds(_meta.HoldToLineMs);
            if (TxtSetHoldMs != null) TxtSetHoldMs.Text = $"{_meta.HoldToLineMs} ms";
            SaveMetadata();
        }

        private void SliderSetSmoothing_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.StylusSmoothing = Math.Round(e.NewValue, 1);
            if (TxtSetSmoothing != null) TxtSetSmoothing.Text = $"{_meta.StylusSmoothing:0.0} px";
            SaveMetadata();
        }

        private void CmbSetBarrelMode_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || CmbSetBarrelMode == null) return;
            _meta.BarrelButtonMode = CmbSetBarrelMode.SelectedIndex;
            SaveMetadata();
        }

        private void ChkSetEraserAura_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || ChkSetEraserAura == null) return;
            _meta.EraserAuraEnabled = ChkSetEraserAura.IsChecked == true;
            if (!_meta.EraserAuraEnabled) ClearEraserHover();
            SaveMetadata();
        }

        private void BtnRenderQuality_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button b || b.Tag is not string s || !double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double sc)) return;
        _meta.RenderScale = sc;
        UpdateQualityCardHighlights();
        SaveMetadata();

        if (_activeSession != null)
        {
            _activeSession.HighResCache.Clear();
            _activeSession.HighResLru.Clear();
            foreach (var slot in _activeSession.Slots)
                slot.HasHighResRendered = false;
            TriggerBackgroundPageRender();
        }
    }

    private void UpdateQualityCardHighlights()
        {
            if (BtnQualBalanced == null) return;
            var actBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252A3D"));
            var normBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D202B"));
            var actBd = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6366F1"));

            foreach (var btn in new[] { BtnQualFast, BtnQualBalanced, BtnQualUltra })
            {
                if (btn?.Tag is string t && double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                {
                    bool sel = Math.Abs(v - _meta.RenderScale) < 0.15;
                    btn.Background = sel ? actBg : normBg;
                    btn.BorderBrush = sel ? actBd : Brushes.Transparent;
                }
            }
        }

        private void SliderSetCachePages_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.MaxCachePages = (int)Math.Round(e.NewValue);
            if (TxtSetCachePages != null) TxtSetCachePages.Text = $"{_meta.MaxCachePages} Sayfa";
            SaveMetadata();
        }

        private void SliderSetScrollSpeed_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.ScrollSpeed = Math.Round(e.NewValue, 2);
            if (TxtSetScrollSpeed != null) TxtSetScrollSpeed.Text = $"{_meta.ScrollSpeed:0.00}x";
            SaveMetadata();
        }

        private void SliderSetDefZoom_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.DefaultZoom = Math.Round(e.NewValue / 100.0, 2);
            if (TxtSetDefZoom != null) TxtSetDefZoom.Text = $"%{(int)Math.Round(e.NewValue)}";
            SaveMetadata();
        }

        private void BtnPageFilter_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button b || b.Tag is not string s || !int.TryParse(s, out int mode)) return;
            _meta.PageFilterMode = mode;
            UpdatePageFilterHighlights();
            ApplyPageFilterEffect();
            SaveMetadata();
        }

        private void UpdatePageFilterHighlights()
        {
            if (BtnFilterNormal == null) return;
            var actBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#252A3D"));
            var normBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D202B"));
            var actBd = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6366F1"));
            var btns = new[] { BtnFilterNormal, BtnFilterSepia, BtnFilterDim };
            for (int i = 0; i < btns.Length; i++)
            {
                if (btns[i] == null) continue;
                bool sel = _meta.PageFilterMode == i;
                btns[i].Background = sel ? actBg : normBg;
                btns[i].BorderBrush = sel ? actBd : Brushes.Transparent;
            }
        }

        private void ApplyPageFilterEffect()
        {
            if (PagesStack == null) return;
            PagesStack.Opacity = _meta.PageFilterMode == 2 ? 0.84 : (_meta.PageFilterMode == 1 ? 0.94 : 1.0);
        }

        private void SliderSetMarginWidth_Changed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _suppressSettingsEvents) return;
            _meta.MarginWidthPx = (int)Math.Round(e.NewValue);
            if (TxtSetMarginWidth != null) TxtSetMarginWidth.Text = $"{_meta.MarginWidthPx} px";
            _cachedMarginBitmap = null;
            if (_activeSession != null && _activeSession.ExtraMarginEnabled)
                MarginOption_Changed(this, new RoutedEventArgs());
            SaveMetadata();
        }

        private void ChkSetStartDrawer_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _suppressSettingsEvents || ChkSetStartDrawer == null) return;
            _meta.StartWithDrawerOpen = ChkSetStartDrawer.IsChecked == true;
            SaveMetadata();
        }

        private void RefreshStorageStatistics()
        {
            try
            {
                int bookCount = _meta.OriginalPaths.Values.Where(File.Exists).Concat(Directory.GetFiles(_libraryFolder, "*.pdf"))
                    .Select(IOPath.GetFileName).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                if (TxtStatBooks != null) TxtStatBooks.Text = $"{bookCount} PDF";

                var isfFiles = Directory.Exists(_annotationsFolder)
                    ? Directory.GetFiles(_annotationsFolder, "*.isf", SearchOption.AllDirectories)
                    : Array.Empty<string>();

                long totalBytes = isfFiles.Sum(f => new FileInfo(f).Length);
                if (TxtStatAnnotatedPages != null) TxtStatAnnotatedPages.Text = $"{isfFiles.Length} Sayfa";
                if (TxtStatDiskSize != null) TxtStatDiskSize.Text = $"{(totalBytes / (1024.0 * 1024.0)):0.00} MB";
            }
            catch { }
        }

        private void BtnBackupAnnotations_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveAllRealizedStrokes();
                SaveMetadata();
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string backupDir = IOPath.Combine(desktop, $"SarpNote_Yedek_{DateTime.Now:yyyy-MM-dd_HH-mm}");
                Directory.CreateDirectory(backupDir);

                if (File.Exists(_metaFilePath))
                    File.Copy(_metaFilePath, IOPath.Combine(backupDir, "metadata.json"), true);

                if (Directory.Exists(_annotationsFolder))
                {
                    foreach (string dirPath in Directory.GetDirectories(_annotationsFolder, "*", SearchOption.AllDirectories))
                        Directory.CreateDirectory(dirPath.Replace(_annotationsFolder, IOPath.Combine(backupDir, "Annotations")));

                    foreach (string newPath in Directory.GetFiles(_annotationsFolder, "*.*", SearchOption.AllDirectories))
                        File.Copy(newPath, newPath.Replace(_annotationsFolder, IOPath.Combine(backupDir, "Annotations")), true);
                }

                MessageBox.Show($"Tüm çözümlerin ve yıldızlı soruların başarıyla yedeklendi:\n{backupDir}", "Yedekleme Tamamlandı", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Yedekleme sırasında hata oluştu: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnClearRamCache_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _tabSessions.Values)
            {
                s.HighResCache.Clear();
                s.HighResLru.Clear();
            }
            _coverCache.Clear();
            TriggerBackgroundPageRender();
            MessageBox.Show("RAM doku ve kapak önbelleği temizlendi.", "SarpNote", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnOpenBaseDir_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start(new ProcessStartInfo { FileName = _baseDir, UseShellExecute = true }); } catch { }
        }

        private void BtnResetSettings_Click(object sender, RoutedEventArgs e)
        {
            _meta.HoldToLineMs = 550;
            _meta.StylusSmoothing = 0.8;
            _meta.BarrelButtonMode = 0;
            _meta.EraserAuraEnabled = true;
            _meta.RenderScale = 2.1;
            _meta.MaxCachePages = 24;
            _meta.ScrollSpeed = 1.05;
            _meta.DefaultZoom = 1.30;
            _meta.PageFilterMode = 0;
            _meta.MarginWidthPx = 560;
            _meta.StartWithDrawerOpen = true;
            _meta.StartupBehavior = 0;
            _meta.StartupPenColorHex = "#0F172A";
            _meta.StartupPenThickness = 1.5;
            _meta.StartupEraserSize = 10.0;
            _meta.StartupExtraMargin = false;
            _meta.StartMaximized = true;
            ApplyLoadedSettingsToRuntime();
            SaveMetadata();
        }
        // #endregion SETTINGS_ENGINE
        protected override void OnClosed(EventArgs e)
        {
            _renderCts.Cancel();
            StopRenderLoop();
            SaveAllRealizedStrokes();
            SaveMetadata();
            foreach (var session in _tabSessions.Values)
            {
                lock (session.SyncLock)
                {
                    session.Document?.Dispose();
                    session.Document = null;
                }
            }
            base.OnClosed(e);
        }
        // #region DUAL_SELECTION_ENGINE
        private string _selectSubMode = "Lasso"; // "Lasso" veya "Rect"
        private bool _isSelectToolExclusiveActive = false;
        private bool _dualSelectionHooked = false;
        private bool _isRectSelecting = false;
        private InkCanvas _rectSelectInk = null;
        private Point _rectSelectStart;
        private RectSelectionAdorner _rectSelectAdorner = null;
        private readonly Dictionary<DependencyObject, (Brush Bg, Brush Border, object Tag)> _savedDockState = new Dictionary<DependencyObject, (Brush, Brush, object)>();

        private sealed class RectSelectionAdorner : System.Windows.Documents.Adorner
        {
            private readonly Point _start;
            private Point _current;
            private static readonly Pen BorderPen;
            private static readonly Brush FillBrush;

            static RectSelectionAdorner()
            {
                BorderPen = new Pen(new SolidColorBrush(Color.FromRgb(129, 140, 248)), 1.6)
                {
                    DashStyle = new DashStyle(new double[] { 4, 3 }, 0),
                    DashCap = PenLineCap.Round
                };
                BorderPen.Freeze();
                FillBrush = new SolidColorBrush(Color.FromArgb(38, 99, 102, 241));
                FillBrush.Freeze();
            }

            public RectSelectionAdorner(UIElement adornedElement, Point start) : base(adornedElement)
            {
                IsHitTestVisible = false;
                _start = _current = start;
            }

            public void UpdatePoint(Point current)
            {
                _current = current;
                InvalidateVisual();
            }

            protected override void OnRender(DrawingContext dc)
            {
                Rect r = new Rect(_start, _current);
                if (r.Width > 1 || r.Height > 1)
                {
                    dc.DrawRoundedRectangle(FillBrush, BorderPen, r, 4, 4);
                }
            }
        }

        private void EnsureDualSelectionEngineHooked()
        {
            if (_dualSelectionHooked) return;
            _dualSelectionHooked = true;
            AddHandler(UIElement.PreviewMouseLeftButtonDownEvent, new MouseButtonEventHandler(DualSelect_PreviewMouseDown), false);
            AddHandler(UIElement.PreviewMouseMoveEvent, new MouseEventHandler(DualSelect_PreviewMouseMove), false);
            AddHandler(UIElement.PreviewMouseLeftButtonUpEvent, new MouseButtonEventHandler(DualSelect_PreviewMouseUp), false);
        }

        private void ActivateExclusiveSelectTool()
        {
            EnsureDualSelectionEngineHooked();
            _isSelectToolExclusiveActive = true;
            
            _activeTool = "Lasso";
            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();

            // Tüm açık sayfalardaki InkCanvas'ları kesin olarak Select moduna al
            foreach (var ic in EnumerateRealizedInkCanvases(PagesStack))
            {
                ic.EditingMode = InkCanvasEditingMode.Select;
            }

            SyncDualSelectButtons();
        }

        private void BtnUnifiedSelect_Click(object sender, RoutedEventArgs e)
        {
            ActivateExclusiveSelectTool();
            if (SelectModePopup != null)
            {
                SelectModePopup.IsOpen = true;
            }
        }

        private void BtnUnifiedSelect_RightClick(object sender, MouseButtonEventArgs e)
        {
            ActivateExclusiveSelectTool();
            if (SelectModePopup != null)
            {
                SelectModePopup.IsOpen = true;
                e.Handled = true;
            }
        }

        private void BtnPickLassoMode_Click(object sender, RoutedEventArgs e)
        {
            _selectSubMode = "Lasso";
            if (SelectModePopup != null) SelectModePopup.IsOpen = false;
            ActivateExclusiveSelectTool();
        }

        private void BtnPickRectMode_Click(object sender, RoutedEventArgs e)
        {
            _selectSubMode = "Rect";
            if (SelectModePopup != null) SelectModePopup.IsOpen = false;
            ActivateExclusiveSelectTool();
        }

        private void RestoreOtherDockButtonsVisuals() => HighlightActiveToolButtons();

        private void ClearOtherDockButtonsVisuals(DependencyObject root) { }

        private void SyncDualSelectButtons()
        {
            bool isSelectActive = _isSelectToolExclusiveActive &&
                (string.Equals(_activeTool, "Lasso", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(_activeTool, "Select", StringComparison.OrdinalIgnoreCase));

            if (BtnLasso != null)
            {
                BtnLasso.ApplyTemplate();
                BtnLasso.Tag = isSelectActive ? "Active" : null;

                if (BtnLasso.Template != null)
                {
                    if (BtnLasso.Template.FindName("IconSelectLasso", BtnLasso) is UIElement icoLasso)
                        icoLasso.Visibility = _selectSubMode == "Rect" ? Visibility.Collapsed : Visibility.Visible;
                    if (BtnLasso.Template.FindName("IconSelectRect", BtnLasso) is UIElement icoRect)
                        icoRect.Visibility = _selectSubMode == "Rect" ? Visibility.Visible : Visibility.Collapsed;
                }

                if (isSelectActive)
                {
                    // Yan dock konteynerini bul ve diğer tüm araçların (El, Kalem 1.5, Silgi vb.) seçili çerçevesini temizle
                    DependencyObject dockRoot = System.Windows.Media.VisualTreeHelper.GetParent(BtnLasso);
                    if (dockRoot != null && System.Windows.Media.VisualTreeHelper.GetParent(dockRoot) is Panel parentPanel)
                        dockRoot = parentPanel;
                    ClearOtherDockButtonsVisuals(dockRoot);
                }
            }

            if (BtnPickLassoMode != null && BtnPickLassoMode.Template != null)
            {
                BtnPickLassoMode.ApplyTemplate();
                if (BtnPickLassoMode.Template.FindName("Row", BtnPickLassoMode) is Border r1)
                {
                    r1.Background = _selectSubMode == "Lasso" ? new SolidColorBrush(Color.FromRgb(30, 36, 59)) : new SolidColorBrush(Color.FromRgb(22, 27, 44));
                    r1.BorderBrush = _selectSubMode == "Lasso" ? new SolidColorBrush(Color.FromRgb(99, 102, 241)) : new SolidColorBrush(Color.FromRgb(37, 44, 68));
                }
            }

            if (BtnPickRectMode != null && BtnPickRectMode.Template != null)
            {
                BtnPickRectMode.ApplyTemplate();
                if (BtnPickRectMode.Template.FindName("Row", BtnPickRectMode) is Border r2)
                {
                    r2.Background = _selectSubMode == "Rect" ? new SolidColorBrush(Color.FromRgb(30, 36, 59)) : new SolidColorBrush(Color.FromRgb(22, 27, 44));
                    r2.BorderBrush = _selectSubMode == "Rect" ? new SolidColorBrush(Color.FromRgb(56, 189, 248)) : new SolidColorBrush(Color.FromRgb(37, 44, 68));
                }
            }
        }

        private static IEnumerable<InkCanvas> EnumerateRealizedInkCanvases(DependencyObject root)
        {
            if (root == null) yield break;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
                if (child is InkCanvas ic) yield return ic;
                foreach (var nested in EnumerateRealizedInkCanvases(child))
                    yield return nested;
            }
        }

        // #region PAGE_SIZE_NORMALIZER
        private static double GetNormalizedPageWidth(System.Drawing.SizeF sz)
        {
            double rawW = Math.Max(1.0, sz.Width);
            double rawH = Math.Max(1.0, sz.Height);
            // %100'de 860px, %130 varsayılan açılışta 1118px (1080p ekranda ideal okuma boyutu)
            return rawH >= rawW ? 860.0 : 1180.0;
        }

        private static double GetNormalizedPageHeight(System.Drawing.SizeF sz)
        {
            double rawW = Math.Max(1.0, sz.Width);
            double rawH = Math.Max(1.0, sz.Height);
            double normW = rawH >= rawW ? 860.0 : 1180.0;
            return Math.Round(normW * (rawH / rawW));
        }

        private void ApplyHighQualityPageRenderingOptions()
        {
            if (PagesStack != null)
            {
                RenderOptions.SetBitmapScalingMode(PagesStack, BitmapScalingMode.HighQuality);
                RenderOptions.SetClearTypeHint(PagesStack, ClearTypeHint.Enabled);
                PagesStack.UseLayoutRounding = true;
                PagesStack.SnapsToDevicePixels = true;
            }
            PdfCrispRenderExtensions.CurrentZoomFactor = _activeSession != null ? _activeSession.ZoomLevel : 1.30;
        }
        // #endregion PAGE_SIZE_NORMALIZER

                        // #region PRECISION_SHAPE_ENGINE
        private bool _precisionShapeHooked = false;
        private System.Windows.Threading.DispatcherTimer _precisionHoldTimer = null;
        private InkCanvas _activeDrawingInk = null;
        private readonly List<(System.Windows.Point Pt, long TimeMs)> _liveRawPoints = new List<(System.Windows.Point, long)>();
        private bool _isPenCurrentlyDown = false;
        private bool _hasSnappedDuringCurrentHold = false;
        private bool _isInternalStrokeReplace = false;
        private System.Windows.Ink.Stroke _liveSnappedStroke = null;
        private string _liveSnappedKind = null;
        private System.Windows.Point _liveSnappedLineStart;

        private static readonly Guid GeoShapeTypeGuid = new Guid("A8C19B41-4F21-47E9-9B8A-112233445566");
        private static readonly Guid GeoVerticesDataGuid = new Guid("B9D28C52-5E32-48FA-AC9B-223344556677");

        private sealed class SnapPulseAdorner : System.Windows.Documents.Adorner
        {
            private readonly List<System.Windows.Point> _points;
            private double _progress = 0.0;
            private readonly System.Windows.Threading.DispatcherTimer _animTimer;
            private readonly System.Windows.Documents.AdornerLayer _layer;

            public SnapPulseAdorner(UIElement adornedElement, System.Windows.Documents.AdornerLayer layer, IEnumerable<System.Windows.Point> pts)
                : base(adornedElement)
            {
                IsHitTestVisible = false;
                _layer = layer;
                _points = pts.Take(8).ToList();
                _animTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
                _animTimer.Tick += (s, e) =>
                {
                    _progress += 0.08;
                    if (_progress >= 1.0)
                    {
                        _animTimer.Stop();
                        try { _layer?.Remove(this); } catch { }
                    }
                    else
                    {
                        InvalidateVisual();
                    }
                };
                _animTimer.Start();
            }

            protected override void OnRender(DrawingContext dc)
            {
                double alpha = Math.Max(0, 1.0 - _progress);
                var pen = new System.Windows.Media.Pen(
                    new SolidColorBrush(System.Windows.Media.Color.FromArgb((byte)(alpha * 235), 56, 189, 248)), 2.2);
                var fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb((byte)(alpha * 110), 99, 102, 241));
                double radius = 4.0 + (_progress * 9.5);

                foreach (var p in _points)
                {
                    dc.DrawEllipse(fill, pen, p, radius, radius);
                }
            }
        }

                

        
        

        
        private bool IsFreehandPenOrHighlighterActive(InkCanvas ink)
        {
            if (ink == null) return false;
            if (_isSelectToolExclusiveActive) return false;

            string t = _activeTool ?? "Pen";
            return !string.Equals(t, "Eraser", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Silgi", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Lasso", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Select", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(t, "Pan", StringComparison.OrdinalIgnoreCase);
        }

        private void EnsurePrecisionShapeEngineHooked() { _precisionShapeHooked = true; }

        

        private void BeginPrecisionTracking(InkCanvas ink, System.Windows.Point pt)
        {
            _activeDrawingInk = ink;
            _isPenCurrentlyDown = true;
            _hasSnappedDuringCurrentHold = false;
            _liveSnappedStroke = null;
            _liveSnappedKind = null;
            _liveRawPoints.Clear();

            long now = Environment.TickCount64;
            _liveRawPoints.Add((pt, now));
            _precisionHoldTimer.Start();
        }

        private void UpdatePrecisionTrackingPoint(System.Windows.Point pt)
        {
            if (!_isPenCurrentlyDown || _activeDrawingInk == null) return;
            long now = Environment.TickCount64;

            // Eğer basılı tutarken şekle/çizgiye dönüştüyse ve kullanıcı kalemi kaldırmadan sürüklüyorsa:
            if (_hasSnappedDuringCurrentHold)
            {
                if (_liveSnappedKind == "Line" && _liveSnappedStroke != null && _activeDrawingInk.Strokes.Contains(_liveSnappedStroke))
                {
                    if (Dist(_liveSnappedLineStart, pt) >= 10.0)
                    {
                        var s0 = _liveSnappedLineStart;
                        var s1 = pt;
                        ApplyMagneticSnapToLine(_activeDrawingInk, ref s0, ref s1, _liveSnappedStroke);
                        _liveSnappedStroke.StylusPoints = CreateDenseLineStylusPoints(s0, s1);
                        SaveStrokeVerticesMetadata(_liveSnappedStroke, "Line", new[] { s0, s1 });
                    }
                }
                return;
            }

            if (_liveRawPoints.Count == 0 || Dist(_liveRawPoints[_liveRawPoints.Count - 1].Pt, pt) >= 0.8)
            {
                _liveRawPoints.Add((pt, now));
            }
            else
            {
                // Aynı noktada dururken zaman damgasını güncel tutmak için son noktanın zamanını bilerek değiştirmiyoruz
            }
        }

        private void EndPrecisionTracking()
        {
            if (_activeDrawingInk != null && _hasSnappedDuringCurrentHold)
            {
                if (_liveSnappedKind == "Line" && _liveSnappedStroke != null && _activeDrawingInk.Strokes.Contains(_liveSnappedStroke))
                {
                    var verts = GetStrokeVertices(_liveSnappedStroke);
                    if (verts != null && verts.Count >= 2)
                    {
                        TryMergeCollinearLines(_activeDrawingInk, _liveSnappedStroke, verts[0], verts[verts.Count - 1]);
                    }
                }

                if (_activeDrawingInk.EditingMode == InkCanvasEditingMode.None)
                {
                    _activeDrawingInk.EditingMode = InkCanvasEditingMode.Ink;
                }
            }

            _precisionHoldTimer?.Stop();
            _isPenCurrentlyDown = false;
        }

        private void PrecisionShape_PreviewStylusDown(object sender, StylusDownEventArgs e)
        {
            var ink = FindParentInkCanvas(e.OriginalSource as DependencyObject);
            if (!IsFreehandPenOrHighlighterActive(ink)) return;
            BeginPrecisionTracking(ink, e.GetPosition(ink));
        }

        private void PrecisionShape_PreviewStylusMove(object sender, StylusEventArgs e)
        {
            if (!_isPenCurrentlyDown || _activeDrawingInk == null) return;
            if (e.InAir) return;
            UpdatePrecisionTrackingPoint(e.GetPosition(_activeDrawingInk));
        }

        private void PrecisionShape_PreviewStylusUp(object sender, StylusEventArgs e)
        {
            EndPrecisionTracking();
        }

        private void PrecisionShape_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return; // Stylus zaten yakaladı
            var ink = FindParentInkCanvas(e.OriginalSource as DependencyObject);
            if (!IsFreehandPenOrHighlighterActive(ink)) return;
            BeginPrecisionTracking(ink, e.GetPosition(ink));
        }

        private void PrecisionShape_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.StylusDevice != null) return;
            if (!_isPenCurrentlyDown || _activeDrawingInk == null) return;
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndPrecisionTracking();
                return;
            }
            UpdatePrecisionTrackingPoint(e.GetPosition(_activeDrawingInk));
        }

        private void PrecisionShape_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return;
            EndPrecisionTracking();
        }

        // Kalem sabit tutulduğunda (280 ms) çalışan canlı dönüştürücü
        private void PrecisionHoldTimer_Tick(object sender, EventArgs e)
        {
            if (!_isPenCurrentlyDown || _hasSnappedDuringCurrentHold || _activeDrawingInk == null)
            {
                _precisionHoldTimer.Stop();
                return;
            }

            if (_liveRawPoints.Count < 6) return;

            long now = Environment.TickCount64;
            var lastEntry = _liveRawPoints[_liveRawPoints.Count - 1];

            // Son 280 ms içinde kalem 9 pikselden küçük bir alanda sabit tutuldu mu? (Tablet titremesini tolere eder)
            long stationarySince = lastEntry.TimeMs;
            for (int i = _liveRawPoints.Count - 1; i >= 0; i--)
            {
                if (Dist(_liveRawPoints[i].Pt, lastEntry.Pt) <= 9.0)
                    stationarySince = _liveRawPoints[i].TimeMs;
                else
                    break;
            }

            long holdDuration = now - stationarySince;
            if (holdDuration < 550) return;

            var cleanedPts = CleanHoldTailAndResample(_liveRawPoints);
            if (cleanedPts.Count < 5 || PolylineLength(cleanedPts) < 24.0) return;

            if (TryBuildPrecisionGeometricStroke(
                _activeDrawingInk,
                cleanedPts,
                out System.Windows.Ink.Stroke geoStroke,
                out string shapeKind,
                out List<System.Windows.Point> keyVertices))
            {
                _hasSnappedDuringCurrentHold = true;
                _liveSnappedStroke = geoStroke;
                _liveSnappedKind = shapeKind;
                if (keyVertices.Count >= 1) _liveSnappedLineStart = keyVertices[0];

                // Islak mürekkebi iptal et ve şekli ekle
                _isInternalStrokeReplace = true;
                try
                {
                    _activeDrawingInk.EditingMode = InkCanvasEditingMode.None;
                    _activeDrawingInk.Strokes.Add(geoStroke);
                }
                finally
                {
                    _isInternalStrokeReplace = false;
                }

                if (shapeKind == "Line" && keyVertices.Count == 2)
                {
                    TryMergeCollinearLines(_activeDrawingInk, geoStroke, keyVertices[0], keyVertices[1]);
                }

                ShowSnapPulse(_activeDrawingInk, keyVertices);
                SaveAllRealizedStrokes();
            }
            else
            {
                _precisionHoldTimer.Stop();
            }
        }

        // Kalem kaldırıldığında (StrokeCollected) çalışan ikinci güvenlik ve otomatik birleştirme katmanı
        private void PrecisionShape_StrokeCollected(object sender, InkCanvasStrokeCollectedEventArgs e)
        {
            if (_isInternalStrokeReplace) return;
            var ink = sender as InkCanvas;
            if (ink == null || e.Stroke == null) return;

            // Eğer basılı tutarken zaten şekle dönüştüyse, InkCanvas'ın sonradan ürettiği ham çizimi sil
            if (_hasSnappedDuringCurrentHold)
            {
                if (ink.Strokes.Contains(e.Stroke) && !ReferenceEquals(e.Stroke, _liveSnappedStroke))
                {
                    ink.Strokes.Remove(e.Stroke);
                }
                _hasSnappedDuringCurrentHold = false;
                return;
            }

            // Kullanıcı basılı tutup tam dönüştürme anında kalemi kaldırdıysa (veya son noktada duraksadıysa)
            long now = Environment.TickCount64;
            bool pausedAtEnd = false;
            if (_liveRawPoints.Count >= 6)
            {
                var lastPt = _liveRawPoints[_liveRawPoints.Count - 1].Pt;
                long stationarySince = _liveRawPoints[_liveRawPoints.Count - 1].TimeMs;
                for (int i = _liveRawPoints.Count - 1; i >= 0; i--)
                {
                    if (Dist(_liveRawPoints[i].Pt, lastPt) <= 10.0)
                        stationarySince = _liveRawPoints[i].TimeMs;
                    else
                        break;
                }
                if (now - stationarySince >= 190) pausedAtEnd = true;
            }

            var rawStrokePts = e.Stroke.StylusPoints.Select(sp => new System.Windows.Point(sp.X, sp.Y)).ToList();
            if (rawStrokePts.Count < 2) return;

            // A) Eğer kullanıcı sonunda kısa süre beklediyse şekil/çizgi dönüşümünü doğrudan e.Stroke üzerinde uygula
            if (pausedAtEnd && _liveRawPoints.Count >= 6)
            {
                var cleanedPts = CleanHoldTailAndResample(_liveRawPoints);
                if (cleanedPts.Count >= 5 &&
                    TryBuildPrecisionGeometricStroke(ink, cleanedPts, out var geoStroke, out string shapeKind, out var keyVertices))
                {
                    e.Stroke.StylusPoints = geoStroke.StylusPoints;
                    e.Stroke.DrawingAttributes.FitToCurve = false;
                    SaveStrokeVerticesMetadata(e.Stroke, shapeKind, keyVertices);

                    if (shapeKind == "Line" && keyVertices.Count == 2)
                    {
                        TryMergeCollinearLines(ink, e.Stroke, keyVertices[0], keyVertices[1]);
                    }

                    ShowSnapPulse(ink, keyVertices);
                    SaveAllRealizedStrokes();
                    return;
                }
            }

            // B) Kullanıcı hiç beklemeden bile çok düzgün bir doğru parçası çizdiyse ve ucu mevcut bir çizginin ucuna yakınsa mıknatısla ve birleştir!
            double len = PolylineLength(rawStrokePts);
            double direct = Dist(rawStrokePts[0], rawStrokePts[rawStrokePts.Count - 1]);
            if (len >= 24.0 &&
                direct / Math.Max(1.0, len) >= 0.968 &&
                MaxPerpendicularDistance(rawStrokePts, rawStrokePts[0], rawStrokePts[rawStrokePts.Count - 1]) <= Math.Max(4.2, direct * 0.032))
            {
                var p0 = rawStrokePts[0];
                var p1 = rawStrokePts[rawStrokePts.Count - 1];
                bool snapped = ApplyMagneticSnapToLine(ink, ref p0, ref p1, e.Stroke);
                e.Stroke.StylusPoints = CreateDenseLineStylusPoints(p0, p1);
                e.Stroke.DrawingAttributes.FitToCurve = false;
                SaveStrokeVerticesMetadata(e.Stroke, "Line", new[] { p0, p1 });
                TryMergeCollinearLines(ink, e.Stroke, p0, p1);
                if (snapped) ShowSnapPulse(ink, new[] { p0, p1 });
                SaveAllRealizedStrokes();
            }
        }

        private static void ShowSnapPulse(InkCanvas ink, IEnumerable<System.Windows.Point> pts)
        {
            // Efekt kaldirildi
        }

        // =====================================================================
        // GEOMETRİK ŞEKİL TANIMA VE DOĞRULAMA MOTORU
        // =====================================================================
        private bool TryBuildPrecisionGeometricStroke(
            InkCanvas ink,
            List<System.Windows.Point> pts,
            out System.Windows.Ink.Stroke resultStroke,
            out string shapeKind,
            out List<System.Windows.Point> keyVertices)
        {
            resultStroke = null;
            shapeKind = null;
            keyVertices = new List<System.Windows.Point>();

            int n = pts.Count;
            if (n < 4) return false;

            var pStart = pts[0];
            var pEnd = pts[n - 1];
            double pathLen = PolylineLength(pts);
            double endDist = Dist(pStart, pEnd);
            System.Windows.Rect bbox = ComputeBounds(pts);
            double diag = Math.Sqrt(bbox.Width * bbox.Width + bbox.Height * bbox.Height);
            if (diag < 7.5) return false;

            // 1. DÜZ ÇİZGİ TESTİ (Küçük ve büyük tüm doğrular için orantılı sapma)
            double directness = endDist / Math.Max(1.0, pathLen);
            double maxDev = MaxPerpendicularDistance(pts, pStart, pEnd);
            double rmsDev = RmsPerpendicularDistance(pts, pStart, pEnd);

            double allowedMaxDev = Math.Clamp(endDist * 0.065, 1.8, 8.5);
            double allowedRmsDev = Math.Clamp(endDist * 0.032, 0.9, 3.8);

            if (endDist >= 7.5 &&
                directness >= 0.900 &&
                maxDev <= allowedMaxDev &&
                rmsDev <= allowedRmsDev)
            {
                var s0 = pStart;
                var s1 = pEnd;
                ApplyMagneticSnapToLine(ink, ref s0, ref s1, null);

                resultStroke = CreateStyledGeometricStroke(ink, CreateDenseLineStylusPoints(s0, s1));
                shapeKind = "Line";
                keyVertices = new List<System.Windows.Point> { s0, s1 };
                SaveStrokeVerticesMetadata(resultStroke, "Line", keyVertices);
                return true;
            }

            // 2. KAPALI ŞEKİL TESTİ (Küçük ve Büyük Daire, Elips, Kare, Dikdörtgen, Üçgen)
            var closedLoop = TryExtractClosedSingleLoop(pts, diag);
            if (closedLoop != null && closedLoop.Count >= 8)
            {
                System.Windows.Rect cBox = ComputeBounds(closedLoop);
                double cDiag = Math.Sqrt(cBox.Width * cBox.Width + cBox.Height * cBox.Height);
                double loopLen = PolylineLength(closedLoop);
                double polyArea = Math.Abs(SignedPolygonArea(closedLoop));
                double compactness = (4.0 * Math.PI * polyArea) / Math.Max(1.0, loopLen * loopLen);

                if (cBox.Width >= 7.5 && cBox.Height >= 7.5 && compactness >= 0.24)
                {
                    var corners = DetectSharpPolygonCorners(closedLoop, cDiag);

                    // A) Daire / Elips Uyumu
                    var center = new System.Windows.Point(cBox.Left + cBox.Width * 0.5, cBox.Top + cBox.Height * 0.5);
                    double rx = cBox.Width * 0.5;
                    double ry = cBox.Height * 0.5;
                    double ellipseRmsError = ComputeNormalizedEllipseError(closedLoop, center, rx, ry);

                    if (corners.Count == 3 || corners.Count == 4)
                    {
                        double edgeStraightErr = ComputePolygonEdgeError(closedLoop, corners, cDiag);
                        if (edgeStraightErr < 0.046)
                            ellipseRmsError += 0.060;
                    }

                    // B) Üçgen Uyumu
                    double triError = double.MaxValue;
                    List<System.Windows.Point> triVerts = null;
                    if (TryFitTriangle(closedLoop, corners, cDiag, out triVerts, out double tErr))
                    {
                        triError = tErr;
                    }

                    // C) Kare / Dikdörtgen Uyumu
                    double rectError = double.MaxValue;
                    List<System.Windows.Point> rectVerts = null;
                    bool isSquare = false;
                    if (TryFitRectangleOrSquare(closedLoop, corners, cBox, cDiag, out rectVerts, out isSquare, out double rErr))
                    {
                        rectError = rErr;
                    }

                    double bestErr = Math.Min(ellipseRmsError, Math.Min(triError, rectError));
                    double maxAllowedShapeErr = cDiag < 32.0 ? 0.092 : 0.078;

                    if (bestErr <= maxAllowedShapeErr)
                    {
                        if (bestErr == triError && triVerts != null)
                        {
                            SnapPolygonVerticesToExisting(ink, triVerts);
                            var closedTri = new List<System.Windows.Point>(triVerts) { triVerts[0] };
                            resultStroke = CreateStyledGeometricStroke(ink, CreateDensePolylineStylusPoints(closedTri));
                            shapeKind = "Triangle";
                            keyVertices = triVerts;
                            SaveStrokeVerticesMetadata(resultStroke, shapeKind, keyVertices);
                            return true;
                        }
                        else if (bestErr == rectError && rectVerts != null)
                        {
                            SnapPolygonVerticesToExisting(ink, rectVerts);
                            var closedRect = new List<System.Windows.Point>(rectVerts) { rectVerts[0] };
                            resultStroke = CreateStyledGeometricStroke(ink, CreateDensePolylineStylusPoints(closedRect));
                            shapeKind = isSquare ? "Square" : "Rectangle";
                            keyVertices = rectVerts;
                            SaveStrokeVerticesMetadata(resultStroke, shapeKind, keyVertices);
                            return true;
                        }
                        else if (bestErr == ellipseRmsError)
                        {
                            double aspect = Math.Max(rx, ry) / Math.Max(1.0, Math.Min(rx, ry));
                            if (aspect <= 1.25)
                            {
                                double r = (rx + ry) * 0.5;
                                rx = ry = r;
                                shapeKind = "Circle";
                            }
                            else
                            {
                                shapeKind = "Ellipse";
                            }

                            resultStroke = CreateStyledGeometricStroke(ink, CreateSmoothEllipseStylusPoints(center, rx, ry));
                            keyVertices = new List<System.Windows.Point>
                            {
                                center,
                                new System.Windows.Point(center.X - rx, center.Y),
                                new System.Windows.Point(center.X + rx, center.Y),
                                new System.Windows.Point(center.X, center.Y - ry),
                                new System.Windows.Point(center.X, center.Y + ry)
                            };
                            SaveStrokeVerticesMetadata(resultStroke, shapeKind, keyVertices);
                            return true;
                        }
                    }
                }
            }

            // 3. AÇIK KIRIK ÇİZGİ / KÜÇÜK DİKLİK SEMBOLÜ / DİK EKSEN
            if (TryFitOpenMultiSegmentPolyline(ink, pts, diag, out List<System.Windows.Point> polyVerts))
            {
                resultStroke = CreateStyledGeometricStroke(ink, CreateDensePolylineStylusPoints(polyVerts));
                shapeKind = "Polyline";
                keyVertices = polyVerts;
                SaveStrokeVerticesMetadata(resultStroke, "Polyline", keyVertices);
                return true;
            }

            return false;
        }

        // =====================================================================
        // MIKNATISLI KÖŞE YAKALAMA VE DOĞRUSAL ÇİZGİ BİRLEŞTİRME MOTORU
        // =====================================================================
        private bool ApplyMagneticSnapToLine(
            InkCanvas ink,
            ref System.Windows.Point pStart,
            ref System.Windows.Point pEnd,
            System.Windows.Ink.Stroke ignoreStroke)
        {
            if (ink == null) return false;
            bool snappedAny = false;
            bool endSnapped = false;

            // Sadece kullanici ucu mevcut bir geometrik koseye cok yakin (9 px) birakirsa kenetle
            const double VertexSnapRadius = 9.0;

            var existingVertices = new List<System.Windows.Point>();
            var existingSegments = new List<(System.Windows.Point A, System.Windows.Point B)>();

            CollectExistingGeometryOnCanvas(ink, ignoreStroke, existingVertices, existingSegments);

            if (TryFindClosestPoint(pStart, existingVertices, VertexSnapRadius, out var snapStartV))
            {
                pStart = snapStartV;
                snappedAny = true;
            }

            if (TryFindClosestPoint(pEnd, existingVertices, VertexSnapRadius, out var snapEndV) && Dist(snapEndV, pStart) > 6.0)
            {
                System.Windows.Point curStart = pStart;
                bool sameExistingSegment = false;
                foreach (var seg in existingSegments)
                {
                    if ((Dist(curStart, seg.A) < 1.5 && Dist(snapEndV, seg.B) < 1.5) ||
                        (Dist(curStart, seg.B) < 1.5 && Dist(snapEndV, seg.A) < 1.5))
                    {
                        sameExistingSegment = true;
                        break;
                    }
                }

                if (!sameExistingSegment)
                {
                    pEnd = snapEndV;
                    endSnapped = true;
                    snappedAny = true;
                }
            }

            // Yatay (0°) ve Dikey (90°) hizalama: Baslangic noktasini (pStart) ASLA kaydirma, sadece serbest bitis ucunu duzelt
            if (!endSnapped)
            {
                double dx = pEnd.X - pStart.X;
                double dy = pEnd.Y - pStart.Y;
                double len = Math.Sqrt(dx * dx + dy * dy);
                if (len >= 7.5)
                {
                    double angleDeg = Math.Atan2(dy, dx) * (180.0 / Math.PI);
                    double[] cardinalAngles = { 0, 90, 180, -90, -180 };
                    foreach (double ca in cardinalAngles)
                    {
                        if (Math.Abs(angleDeg - ca) <= 2.8)
                        {
                            double rad = ca * (Math.PI / 180.0);
                            pEnd = new System.Windows.Point(pStart.X + len * Math.Cos(rad), pStart.Y + len * Math.Sin(rad));
                            snappedAny = true;
                            break;
                        }
                    }
                }
            }

            return snappedAny;
        }

        private void TryMergeCollinearLines(
            InkCanvas ink,
            System.Windows.Ink.Stroke currentLineStroke,
            System.Windows.Point a1,
            System.Windows.Point a2)
        {
            if (ink == null || currentLineStroke == null) return;
            if (Dist(a1, a2) < 18.0) return;

            var toRemove = new List<System.Windows.Ink.Stroke>();
            System.Windows.Point mergedStart = a1;
            System.Windows.Point mergedEnd = a2;

            foreach (var st in ink.Strokes.ToList())
            {
                if (ReferenceEquals(st, currentLineStroke)) continue;
                if (st.DrawingAttributes.Color != currentLineStroke.DrawingAttributes.Color) continue;
                if (Math.Abs(st.DrawingAttributes.Width - currentLineStroke.DrawingAttributes.Width) > 1.8) continue;

                // SADECE geometrik olarak olusturulmus cizgileri birlestirmeye aday kabul et (El yazilarini asla birlestirme)
                if (!TryGetLineEndpoints(st, out var b1, out var b2)) continue;

                // Iki cizginin hangi iki ucunun birbirine temas ettigini (<= 7.5 px) bul
                double d_a1_b1 = Dist(a1, b1);
                double d_a1_b2 = Dist(a1, b2);
                double d_a2_b1 = Dist(a2, b1);
                double d_a2_b2 = Dist(a2, b2);

                double minGap = Math.Min(Math.Min(d_a1_b1, d_a1_b2), Math.Min(d_a2_b1, d_a2_b2));
                if (minGap > 7.5) continue;

                System.Windows.Point jointA, outerA, jointB, outerB;
                if (minGap == d_a1_b1) { jointA = a1; outerA = a2; jointB = b1; outerB = b2; }
                else if (minGap == d_a1_b2) { jointA = a1; outerA = a2; jointB = b2; outerB = b1; }
                else if (minGap == d_a2_b1) { jointA = a2; outerA = a1; jointB = b1; outerB = b2; }
                else { jointA = a2; outerA = a1; jointB = b2; outerB = b1; }

                // KRITIK KONTROL: Iki cizgi temas noktasindan ZIT YONLERE (172° - 180° dogrusal uzanti olarak) gitmeli!
                // Paralel veya yan yana cizgiler ayni yone gittigi icin (dot > 0) burada %100 elenir.
                var jointMid = new System.Windows.Point((jointA.X + jointB.X) * 0.5, (jointA.Y + jointB.Y) * 0.5);
                double vAx = outerA.X - jointMid.X, vAy = outerA.Y - jointMid.Y;
                double vBx = outerB.X - jointMid.X, vBy = outerB.Y - jointMid.Y;
                double lenA = Math.Sqrt(vAx * vAx + vAy * vAy);
                double lenB = Math.Sqrt(vBx * vBx + vBy * vBy);
                if (lenA < 14.0 || lenB < 14.0) continue;

                double cosAngle = (vAx * vBx + vAy * vBy) / (lenA * lenB);
                // cos(173°) = -0.9925 -> Sadece gercek uc uca dogrusal uzantilari birlestir
                if (cosAngle > -0.991) continue;

                // Birlesik dogrudan dikey sapma en fazla 2.5 px olmali
                double perpA = DistancePointToInfiniteLine(jointA, outerA, outerB);
                double perpB = DistancePointToInfiniteLine(jointB, outerA, outerB);
                if (perpA <= 2.5 && perpB <= 2.5)
                {
                    toRemove.Add(st);
                    mergedStart = outerB;
                    mergedEnd = outerA;
                    a1 = mergedStart;
                    a2 = mergedEnd;
                }
            }

            if (toRemove.Count > 0)
            {
                foreach (var old in toRemove)
                {
                    ink.Strokes.Remove(old);
                }

                currentLineStroke.StylusPoints = CreateDenseLineStylusPoints(mergedStart, mergedEnd);
                SaveStrokeVerticesMetadata(currentLineStroke, "Line", new[] { mergedStart, mergedEnd });
            }
        }

        private void SnapPolygonVerticesToExisting(InkCanvas ink, List<System.Windows.Point> polyVerts)
        {
            if (ink == null || polyVerts == null) return;
            var existingVertices = new List<System.Windows.Point>();
            var existingSegments = new List<(System.Windows.Point A, System.Windows.Point B)>();
            CollectExistingGeometryOnCanvas(ink, null, existingVertices, existingSegments);

            for (int i = 0; i < polyVerts.Count; i++)
            {
                if (TryFindClosestPoint(polyVerts[i], existingVertices, 9.0, out var snapped))
                {
                    polyVerts[i] = snapped;
                }
            }
        }

        private void CollectExistingGeometryOnCanvas(
            InkCanvas ink,
            System.Windows.Ink.Stroke ignoreStroke,
            List<System.Windows.Point> outVertices,
            List<(System.Windows.Point A, System.Windows.Point B)> outSegments)
        {
            foreach (var st in ink.Strokes)
            {
                if (ReferenceEquals(st, ignoreStroke)) continue;
                if (!st.ContainsPropertyData(GeoShapeTypeGuid)) continue;

                var metaVerts = GetStrokeVertices(st);
                if (metaVerts != null && metaVerts.Count >= 2)
                {
                    string kind = st.GetPropertyData(GeoShapeTypeGuid) as string ?? "Line";
                    if (kind == "Circle" || kind == "Ellipse") continue;

                    foreach (var v in metaVerts) outVertices.Add(v);

                    if (kind == "Line" || kind == "Polyline")
                    {
                        for (int i = 0; i < metaVerts.Count - 1; i++)
                            outSegments.Add((metaVerts[i], metaVerts[i + 1]));
                    }
                    else if (kind == "Triangle" || kind == "Rectangle" || kind == "Square")
                    {
                        for (int i = 0; i < metaVerts.Count; i++)
                            outSegments.Add((metaVerts[i], metaVerts[(i + 1) % metaVerts.Count]));
                    }
                }
            }
        }

        private static bool TryGetLineEndpoints(System.Windows.Ink.Stroke st, out System.Windows.Point p0, out System.Windows.Point p1)
        {
            p0 = default;
            p1 = default;
            if (st == null || !st.ContainsPropertyData(GeoShapeTypeGuid)) return false;

            string k = st.GetPropertyData(GeoShapeTypeGuid) as string;
            if (k != "Line") return false;

            var v = GetStrokeVertices(st);
            if (v != null && v.Count == 2)
            {
                p0 = v[0];
                p1 = v[1];
                return Dist(p0, p1) >= 7.5;
            }
            return false;
        }

        private static void SaveStrokeVerticesMetadata(System.Windows.Ink.Stroke st, string shapeKind, IEnumerable<System.Windows.Point> verts)
        {
            if (st == null || verts == null) return;
            try
            {
                if (st.ContainsPropertyData(GeoShapeTypeGuid))
                    st.RemovePropertyData(GeoShapeTypeGuid);
                if (st.ContainsPropertyData(GeoVerticesDataGuid))
                    st.RemovePropertyData(GeoVerticesDataGuid);

                st.AddPropertyData(GeoShapeTypeGuid, shapeKind);
                string data = string.Join(";", verts.Select(v =>
                    string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:F2},{1:F2}", v.X, v.Y)));
                st.AddPropertyData(GeoVerticesDataGuid, data);
            }
            catch { }
        }

        private static List<System.Windows.Point> GetStrokeVertices(System.Windows.Ink.Stroke st)
        {
            if (st == null || !st.ContainsPropertyData(GeoVerticesDataGuid)) return null;
            try
            {
                string raw = st.GetPropertyData(GeoVerticesDataGuid) as string;
                if (string.IsNullOrWhiteSpace(raw)) return null;
                var list = new List<System.Windows.Point>();
                foreach (var part in raw.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var xy = part.Split(',');
                    if (xy.Length == 2 &&
                        double.TryParse(xy[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double x) &&
                        double.TryParse(xy[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double y))
                    {
                        list.Add(new System.Windows.Point(x, y));
                    }
                }
                return list;
            }
            catch
            {
                return null;
            }
        }

        // =====================================================================
        // ALT GEOMETRİK HESAPLAMA VE FIT FONKSİYONLARI
        // =====================================================================
        private static List<System.Windows.Point> CleanHoldTailAndResample(List<(System.Windows.Point Pt, long TimeMs)> raw)
        {
            var pts = raw.Select(r => r.Pt).ToList();
            if (pts.Count < 4) return pts;

            var endPt = pts[pts.Count - 1];
            int trimEnd = pts.Count - 1;
            while (trimEnd > 2 && Dist(pts[trimEnd - 1], endPt) <= 6.5)
            {
                trimEnd--;
            }
            pts = pts.Take(trimEnd + 1).ToList();

            var resampled = new List<System.Windows.Point> { pts[0] };
            double step = 2.5;
            double accum = 0.0;

            for (int i = 1; i < pts.Count; i++)
            {
                var p0 = pts[i - 1];
                var p1 = pts[i];
                double seg = Dist(p0, p1);
                if (seg < 0.001) continue;

                while (accum + seg >= step)
                {
                    double t = (step - accum) / seg;
                    var interp = new System.Windows.Point(p0.X + t * (p1.X - p0.X), p0.Y + t * (p1.Y - p0.Y));
                    resampled.Add(interp);
                    p0 = interp;
                    seg = Dist(p0, p1);
                    accum = 0.0;
                }
                accum += seg;
            }

            if (Dist(resampled[resampled.Count - 1], pts[pts.Count - 1]) > 0.5)
                resampled.Add(pts[pts.Count - 1]);

            return resampled;
        }

        private static List<System.Windows.Point> TryExtractClosedSingleLoop(List<System.Windows.Point> pts, double diag)
        {
            int n = pts.Count;
            if (n < 8) return null;

            // Küçük şekillerde kapanış eşiğini şeklin kendi boyutuna (diag * 0.35) göre ölçekle
            double closeThreshold = Math.Clamp(diag * 0.35, 5.5, 32.0);
            double directEndDist = Dist(pts[0], pts[n - 1]);

            int bestStartIdx = 0;
            int bestEndIdx = n - 1;
            double bestDist = directEndDist;

            int maxStartSearch = Math.Max(2, (int)(n * 0.22));
            int minEndSearch = Math.Min(n - 2, (int)(n * 0.76));

            for (int i = 0; i <= maxStartSearch; i++)
            {
                for (int j = minEndSearch; j < n; j++)
                {
                    double d = Dist(pts[i], pts[j]);
                    if (d < bestDist)
                    {
                        bestDist = d;
                        bestStartIdx = i;
                        bestEndIdx = j;
                    }
                }
            }

            if (bestDist > closeThreshold) return null;

            var loop = pts.Skip(bestStartIdx).Take(bestEndIdx - bestStartIdx + 1).ToList();
            if (loop.Count < 7) return null;

            double cx = loop.Average(p => p.X);
            double cy = loop.Average(p => p.Y);
            double totalAngle = 0.0;
            for (int i = 0; i < loop.Count; i++)
            {
                var p0 = loop[i];
                var p1 = loop[(i + 1) % loop.Count];
                double a0 = Math.Atan2(p0.Y - cy, p0.X - cx);
                double a1 = Math.Atan2(p1.Y - cy, p1.X - cx);
                double da = a1 - a0;
                while (da > Math.PI) da -= 2.0 * Math.PI;
                while (da < -Math.PI) da += 2.0 * Math.PI;
                totalAngle += da;
            }

            double turns = Math.Abs(totalAngle) / (2.0 * Math.PI);
            if (turns < 0.65 || turns > 1.42) return null;

            return loop;
        }

        private static List<System.Windows.Point> DetectSharpPolygonCorners(List<System.Windows.Point> closedLoop, double diag)
        {
            var loopWithClose = new List<System.Windows.Point>(closedLoop) { closedLoop[0] };
            double rdpEps = Math.Clamp(diag * 0.055, 1.3, 8.5);
            var simplified = RamerDouglasPeucker(loopWithClose, rdpEps);

            double closeMergeDist = Math.Clamp(diag * 0.20, 2.5, 20.0);
            if (simplified.Count > 1 && Dist(simplified[0], simplified[simplified.Count - 1]) <= closeMergeDist)
            {
                simplified.RemoveAt(simplified.Count - 1);
            }

            double minCornerSpacing = Math.Clamp(diag * 0.16, 2.8, 16.0);
            for (int i = simplified.Count - 1; i >= 0 && simplified.Count > 3; i--)
            {
                int next = (i + 1) % simplified.Count;
                if (Dist(simplified[i], simplified[next]) < minCornerSpacing)
                {
                    simplified.RemoveAt(next);
                }
            }

            for (int i = simplified.Count - 1; i >= 0 && simplified.Count > 3; i--)
            {
                var prev = simplified[(i - 1 + simplified.Count) % simplified.Count];
                var curr = simplified[i];
                var next = simplified[(i + 1) % simplified.Count];
                double interiorDeg = InteriorAngleDegrees(prev, curr, next);
                if (interiorDeg > 152.0)
                {
                    simplified.RemoveAt(i);
                }
            }

            return simplified;
        }

        private static bool TryFitTriangle(
            List<System.Windows.Point> loop,
            List<System.Windows.Point> corners,
            double diag,
            out List<System.Windows.Point> triVerts,
            out double normalizedErr)
        {
            triVerts = null;
            normalizedErr = double.MaxValue;
            if (corners == null || corners.Count != 3) return false;

            var v = new List<System.Windows.Point>(corners);

            for (int i = 0; i < 3; i++)
            {
                double ang = InteriorAngleDegrees(v[(i + 2) % 3], v[i], v[(i + 1) % 3]);
                if (ang < 15.0 || ang > 154.0) return false;
            }

            for (int i = 0; i < 3; i++)
            {
                var a = v[i];
                var b = v[(i + 1) % 3];
                double ang = Math.Atan2(b.Y - a.Y, b.X - a.X) * (180.0 / Math.PI);
                if (Math.Abs(ang) <= 5.5 || Math.Abs(Math.Abs(ang) - 180.0) <= 5.5)
                {
                    double avgY = (a.Y + b.Y) * 0.5;
                    v[i] = new System.Windows.Point(a.X, avgY);
                    v[(i + 1) % 3] = new System.Windows.Point(b.X, avgY);
                }
                else if (Math.Abs(Math.Abs(ang) - 90.0) <= 5.5)
                {
                    double avgX = (a.X + b.X) * 0.5;
                    v[i] = new System.Windows.Point(avgX, a.Y);
                    v[(i + 1) % 3] = new System.Windows.Point(avgX, b.Y);
                }
            }

            double err = ComputePolygonEdgeError(loop, v, diag);
            double maxTriErr = diag < 32.0 ? 0.082 : 0.068;
            if (err > maxTriErr) return false;

            triVerts = v;
            normalizedErr = err;
            return true;
        }

        private static bool TryFitRectangleOrSquare(
            List<System.Windows.Point> loop,
            List<System.Windows.Point> corners,
            System.Windows.Rect cBox,
            double diag,
            out List<System.Windows.Point> rectVerts,
            out bool isSquare,
            out double normalizedErr)
        {
            rectVerts = null;
            isSquare = false;
            normalizedErr = double.MaxValue;

            if (corners == null || corners.Count < 4 || corners.Count > 5) return false;

            double left = cBox.Left, right = cBox.Right, top = cBox.Top, bottom = cBox.Bottom;
            double w = cBox.Width, h = cBox.Height;

            if (Math.Max(w, h) / Math.Max(1.0, Math.Min(w, h)) <= 1.20)
            {
                double side = (w + h) * 0.5;
                double cx = left + w * 0.5;
                double cy = top + h * 0.5;
                left = cx - side * 0.5;
                right = cx + side * 0.5;
                top = cy - side * 0.5;
                bottom = cy + side * 0.5;
                isSquare = true;
            }

            var boxVerts = new List<System.Windows.Point>
            {
                new System.Windows.Point(left, top),
                new System.Windows.Point(right, top),
                new System.Windows.Point(right, bottom),
                new System.Windows.Point(left, bottom)
            };

            double boxErr = ComputePolygonEdgeError(loop, boxVerts, diag);

            if (corners.Count == 4)
            {
                for (int i = 0; i < 4; i++)
                {
                    double ang = InteriorAngleDegrees(corners[(i + 3) % 4], corners[i], corners[(i + 1) % 4]);
                    if (ang < 52.0 || ang > 130.0) return false;
                }
            }

            double maxRectErr = diag < 32.0 ? 0.086 : 0.072;
            if (boxErr <= maxRectErr)
            {
                rectVerts = boxVerts;
                normalizedErr = boxErr;
                return true;
            }

            return false;
        }

        private bool TryFitOpenMultiSegmentPolyline(
            InkCanvas ink,
            List<System.Windows.Point> pts,
            double diag,
            out List<System.Windows.Point> polyVerts)
        {
            polyVerts = null;
            double rdpEps = Math.Clamp(diag * 0.052, 1.4, 8.0);
            var simp = RamerDouglasPeucker(pts, rdpEps);
            if (simp.Count < 3 || simp.Count > 4) return false;

            int startIdx = 0;
            for (int seg = 0; seg < simp.Count - 1; seg++)
            {
                var a = simp[seg];
                var b = simp[seg + 1];
                double segLen = Dist(a, b);
                if (segLen < 7.0) return false;

                int endIdx = pts.Count - 1;
                double bestD = double.MaxValue;
                for (int k = startIdx; k < pts.Count; k++)
                {
                    double d = Dist(pts[k], b);
                    if (d < bestD) { bestD = d; endIdx = k; }
                }

                if (endIdx - startIdx < 2) return false;
                var subPts = pts.Skip(startIdx).Take(endIdx - startIdx + 1).ToList();
                double subDirect = Dist(subPts[0], subPts[subPts.Count - 1]) / Math.Max(1.0, PolylineLength(subPts));
                double subMaxDev = MaxPerpendicularDistance(subPts, a, b);

                if (subDirect < 0.895 || subMaxDev > Math.Clamp(segLen * 0.075, 1.8, 7.5))
                    return false;

                startIdx = endIdx;
            }

            for (int i = 0; i < simp.Count - 1; i++)
            {
                var p0 = simp[i];
                var p1 = simp[i + 1];
                ApplyMagneticSnapToLine(ink, ref p0, ref p1, null);
                simp[i] = p0;
                simp[i + 1] = p1;
            }

            polyVerts = simp;
            return true;
        }

        private static double ComputeNormalizedEllipseError(List<System.Windows.Point> loop, System.Windows.Point center, double rx, double ry)
        {
            if (rx < 3.2 || ry < 3.2) return double.MaxValue;
            double sumSq = 0.0;
            foreach (var p in loop)
            {
                double nx = (p.X - center.X) / rx;
                double ny = (p.Y - center.Y) / ry;
                double rNorm = Math.Sqrt(nx * nx + ny * ny);
                double diff = rNorm - 1.0;
                sumSq += diff * diff;
            }
            return Math.Sqrt(sumSq / loop.Count) * 0.65;
        }

        private static double ComputePolygonEdgeError(List<System.Windows.Point> loop, List<System.Windows.Point> polyVerts, double diag)
        {
            if (polyVerts == null || polyVerts.Count < 3) return double.MaxValue;
            int m = polyVerts.Count;
            double sumSq = 0.0;

            foreach (var p in loop)
            {
                double minD = double.MaxValue;
                for (int i = 0; i < m; i++)
                {
                    double d = DistancePointToSegment(p, polyVerts[i], polyVerts[(i + 1) % m]);
                    if (d < minD) minD = d;
                }
                sumSq += minD * minD;
            }

            return Math.Sqrt(sumSq / loop.Count) / Math.Max(1.0, diag);
        }

                private System.Windows.Ink.Stroke CreateStyledGeometricStroke(InkCanvas ink, StylusPointCollection points)
        {
            bool hi = _activeTool == "Highlighter";
            double calSize = GetCalibratedStrokeThickness();
            var attr = new DrawingAttributes
            {
                Color = GetEffectiveInkColor(),
                Width = hi ? Math.Max(4.5, calSize * 0.34) : calSize,
                Height = hi ? Math.Max(10.0, calSize * 1.18) : calSize,
                StylusTip = hi ? StylusTip.Rectangle : StylusTip.Ellipse,
                IsHighlighter = hi,
                FitToCurve = false,
                IgnorePressure = true
            };
            return new System.Windows.Ink.Stroke(points) { DrawingAttributes = attr };
        }

        private static List<System.Windows.Point> CleanHoldTailAndResampleFromPoints(List<System.Windows.Point> rawPts)
        {
            if (rawPts == null || rawPts.Count < 3)
                return rawPts != null ? new List<System.Windows.Point>(rawPts) : new List<System.Windows.Point>();

            var pts = new List<System.Windows.Point>(rawPts);
            System.Windows.Rect rawBox = ComputeBounds(pts);
            double boxDiag = Math.Sqrt(rawBox.Width * rawBox.Width + rawBox.Height * rawBox.Height);

            // Küçük şekillerde kuyruk temizleme yarıçapını küçült ve şeklin %30'undan fazlasını asla silme
            double trimRadius = Math.Clamp(boxDiag * 0.08, 1.5, 5.5);
            var endPt = pts[pts.Count - 1];
            int trimEnd = pts.Count - 1;
            int minKeepIdx = Math.Max(2, (int)(pts.Count * 0.70));
            while (trimEnd > minKeepIdx && Dist(pts[trimEnd - 1], endPt) <= trimRadius)
            {
                trimEnd--;
            }
            pts = pts.Take(trimEnd + 1).ToList();

            // Küçük şekillerde 1.0 px, büyük şekillerde 2.2 px adım ile yüksek çözünürlüklü örnekleme yap
            double step = boxDiag < 48.0 ? 1.0 : 2.2;
            var resampled = new List<System.Windows.Point> { pts[0] };
            double accum = 0.0;

            for (int i = 1; i < pts.Count; i++)
            {
                var p0 = pts[i - 1];
                var p1 = pts[i];
                double seg = Dist(p0, p1);
                if (seg < 0.001) continue;

                while (accum + seg >= step)
                {
                    double t = (step - accum) / seg;
                    var interp = new System.Windows.Point(p0.X + t * (p1.X - p0.X), p0.Y + t * (p1.Y - p0.Y));
                    resampled.Add(interp);
                    p0 = interp;
                    seg = Dist(p0, p1);
                    accum = 0.0;
                }
                accum += seg;
            }

            if (Dist(resampled[resampled.Count - 1], pts[pts.Count - 1]) > 0.3)
                resampled.Add(pts[pts.Count - 1]);

            return resampled;
        }

        private static StylusPointCollection CreateDenseLineStylusPoints(System.Windows.Point a, System.Windows.Point b)
        {
            var col = new StylusPointCollection();
            double dist = Dist(a, b);
            int steps = Math.Max(2, (int)Math.Ceiling(dist / 6.0));
            for (int i = 0; i <= steps; i++)
            {
                double t = (double)i / steps;
                col.Add(new StylusPoint(a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y), 0.5f));
            }
            return col;
        }

        private static StylusPointCollection CreateDensePolylineStylusPoints(List<System.Windows.Point> verts)
        {
            var col = new StylusPointCollection();
            for (int seg = 0; seg < verts.Count - 1; seg++)
            {
                var a = verts[seg];
                var b = verts[seg + 1];
                double dist = Dist(a, b);
                int steps = Math.Max(2, (int)Math.Ceiling(dist / 6.0));
                int startStep = (seg == 0) ? 0 : 1;
                for (int i = startStep; i <= steps; i++)
                {
                    double t = (double)i / steps;
                    col.Add(new StylusPoint(a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y), 0.5f));
                }
            }
            return col;
        }

        private static StylusPointCollection CreateSmoothEllipseStylusPoints(System.Windows.Point center, double rx, double ry)
        {
            var col = new StylusPointCollection();
            const int segments = 96;
            for (int i = 0; i <= segments; i++)
            {
                double angle = (2.0 * Math.PI * i) / segments;
                double x = center.X + rx * Math.Cos(angle);
                double y = center.Y + ry * Math.Sin(angle);
                col.Add(new StylusPoint(x, y, 0.5f));
            }
            return col;
        }

        private static double Dist(System.Windows.Point a, System.Windows.Point b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static double PolylineLength(List<System.Windows.Point> pts)
        {
            double sum = 0.0;
            for (int i = 1; i < pts.Count; i++) sum += Dist(pts[i - 1], pts[i]);
            return sum;
        }

        private static System.Windows.Rect ComputeBounds(List<System.Windows.Point> pts)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            foreach (var p in pts)
            {
                if (p.X < minX) minX = p.X;
                if (p.Y < minY) minY = p.Y;
                if (p.X > maxX) maxX = p.X;
                if (p.Y > maxY) maxY = p.Y;
            }
            return new System.Windows.Rect(new System.Windows.Point(minX, minY), new System.Windows.Point(maxX, maxY));
        }

        private static double SignedPolygonArea(List<System.Windows.Point> pts)
        {
            double area = 0.0;
            for (int i = 0; i < pts.Count; i++)
            {
                var p0 = pts[i];
                var p1 = pts[(i + 1) % pts.Count];
                area += (p0.X * p1.Y) - (p1.X * p0.Y);
            }
            return area * 0.5;
        }

        private static double DistancePointToInfiniteLine(System.Windows.Point p, System.Windows.Point a, System.Windows.Point b)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y;
            double len = Math.Sqrt(dx * dx + dy * dy);
            if (len < 0.0001) return Dist(p, a);
            return Math.Abs(dy * p.X - dx * p.Y + b.X * a.Y - b.Y * a.X) / len;
        }

        private static double DistancePointToSegment(System.Windows.Point p, System.Windows.Point a, System.Windows.Point b)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y;
            double lenSq = dx * dx + dy * dy;
            if (lenSq < 0.0001) return Dist(p, a);
            double t = ((p.X - a.X) * dx + (p.Y - a.Y) * dy) / lenSq;
            t = Math.Clamp(t, 0.0, 1.0);
            var proj = new System.Windows.Point(a.X + t * dx, a.Y + t * dy);
            return Dist(p, proj);
        }

        private static System.Windows.Point ProjectPointToInfiniteLine(System.Windows.Point p, System.Windows.Point a, System.Windows.Point b)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y;
            double lenSq = dx * dx + dy * dy;
            if (lenSq < 0.0001) return a;
            double t = ((p.X - a.X) * dx + (p.Y - a.Y) * dy) / lenSq;
            return new System.Windows.Point(a.X + t * dx, a.Y + t * dy);
        }

        private static bool IsPointOnSegment(System.Windows.Point p, System.Windows.Point a, System.Windows.Point b)
        {
            double len = Dist(a, b);
            return Dist(p, a) <= len + 1.0 && Dist(p, b) <= len + 1.0;
        }

        private static double MaxPerpendicularDistance(List<System.Windows.Point> pts, System.Windows.Point a, System.Windows.Point b)
        {
            double maxD = 0.0;
            foreach (var p in pts)
            {
                double d = DistancePointToSegment(p, a, b);
                if (d > maxD) maxD = d;
            }
            return maxD;
        }

        private static double RmsPerpendicularDistance(List<System.Windows.Point> pts, System.Windows.Point a, System.Windows.Point b)
        {
            if (pts.Count == 0) return 0.0;
            double sumSq = 0.0;
            foreach (var p in pts)
            {
                double d = DistancePointToSegment(p, a, b);
                sumSq += d * d;
            }
            return Math.Sqrt(sumSq / pts.Count);
        }

        private static double InteriorAngleDegrees(System.Windows.Point prev, System.Windows.Point curr, System.Windows.Point next)
        {
            double v1x = prev.X - curr.X, v1y = prev.Y - curr.Y;
            double v2x = next.X - curr.X, v2y = next.Y - curr.Y;
            double m1 = Math.Sqrt(v1x * v1x + v1y * v1y);
            double m2 = Math.Sqrt(v2x * v2x + v2y * v2y);
            if (m1 < 0.001 || m2 < 0.001) return 180.0;
            double cos = Math.Clamp((v1x * v2x + v1y * v2y) / (m1 * m2), -1.0, 1.0);
            return Math.Acos(cos) * (180.0 / Math.PI);
        }

        private static List<System.Windows.Point> RamerDouglasPeucker(List<System.Windows.Point> pts, double epsilon)
        {
            if (pts == null || pts.Count < 3) return pts != null ? new List<System.Windows.Point>(pts) : new List<System.Windows.Point>();
            int maxIdx = 0;
            double maxDist = 0.0;
            var first = pts[0];
            var last = pts[pts.Count - 1];

            for (int i = 1; i < pts.Count - 1; i++)
            {
                double d = DistancePointToSegment(pts[i], first, last);
                if (d > maxDist)
                {
                    maxDist = d;
                    maxIdx = i;
                }
            }

            if (maxDist > epsilon)
            {
                var left = RamerDouglasPeucker(pts.Take(maxIdx + 1).ToList(), epsilon);
                var right = RamerDouglasPeucker(pts.Skip(maxIdx).ToList(), epsilon);
                left.RemoveAt(left.Count - 1);
                left.AddRange(right);
                return left;
            }

            return new List<System.Windows.Point> { first, last };
        }

        private static bool TryFindClosestPoint(System.Windows.Point target, List<System.Windows.Point> candidates, double radius, out System.Windows.Point best)
        {
            best = target;
            double bestD = radius;
            bool found = false;
            foreach (var c in candidates)
            {
                double d = Dist(target, c);
                if (d <= bestD)
                {
                    bestD = d;
                    best = c;
                    found = true;
                }
            }
            return found;
        }

        private static bool TrySnapPointToExistingSegments(
            System.Windows.Point target,
            List<(System.Windows.Point A, System.Windows.Point B)> segments,
            double radius,
            out System.Windows.Point snappedPt,
            out (System.Windows.Point A, System.Windows.Point B) matchedSeg)
        {
            snappedPt = target;
            matchedSeg = default;
            double bestD = radius;
            bool found = false;

            foreach (var seg in segments)
            {
                var mid = new System.Windows.Point((seg.A.X + seg.B.X) * 0.5, (seg.A.Y + seg.B.Y) * 0.5);
                if (Dist(target, mid) <= 12.0)
                {
                    snappedPt = mid;
                    matchedSeg = seg;
                    return true;
                }

                double dx = seg.B.X - seg.A.X, dy = seg.B.Y - seg.A.Y;
                double lenSq = dx * dx + dy * dy;
                if (lenSq < 1.0) continue;
                double t = Math.Clamp(((target.X - seg.A.X) * dx + (target.Y - seg.A.Y) * dy) / lenSq, 0.0, 1.0);
                var proj = new System.Windows.Point(seg.A.X + t * dx, seg.A.Y + t * dy);
                double d = Dist(target, proj);
                if (d <= bestD)
                {
                    bestD = d;
                    snappedPt = proj;
                    matchedSeg = seg;
                    found = true;
                }
            }

            return found;
        }
        // #endregion PRECISION_SHAPE_ENGINE

                // #region SETTINGS_LAYOUT_FIXER
        private bool _settingsLayoutNormalized = false;

        private void NormalizeSettingsOverlayLayout()
        {
            if (_settingsLayoutNormalized || SettingsOverlay == null) return;
            _settingsLayoutNormalized = true;

            try
            {
                var pages = new FrameworkElement[] { SetPagePen, SetPagePerf, SetPageWork, SetPageStorage, SetPageStartup };
                var valueBadges = new HashSet<TextBlock>
                {
                    TxtSetHoldMs, TxtSetSmoothing, TxtSetCachePages, TxtSetScrollSpeed,
                    TxtSetDefZoom, TxtSetMarginWidth, TxtSetStartPenSize, TxtSetStartEraserSize,
                    TxtStatBooks, TxtStatAnnotatedPages, TxtStatDiskSize
                };

                if (TxtSettingsPageSub != null)
                {
                    TxtSettingsPageSub.TextWrapping = TextWrapping.Wrap;
                }

                // Sag icerik alaninin ScrollViewer icinde oldugundan ve dikey tasmadigindan emin ol
                if (SetPagePen?.Parent is FrameworkElement pageHost && pageHost.Parent is ScrollViewer sv)
                {
                    sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                    sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                    sv.Padding = new Thickness(0, 0, 8, 14);
                }

                foreach (var page in pages)
                {
                    if (page == null) continue;
                    NormalizeSettingsElementRecursive(page, valueBadges);
                }

                // Secim kartlarinin (Basic/Pro, Hizli/Dengeli/Ultra, Normal/Sepya/Gece) metinlerini ortala ve esneklestir
                var modeButtons = new Button[]
                {
                    BtnSetPenBasic, BtnSetPenPro,
                    BtnQualFast, BtnQualBalanced, BtnQualUltra,
                    BtnFilterNormal, BtnFilterSepia, BtnFilterDim
                };

                foreach (var btn in modeButtons)
                {
                    if (btn == null) continue;
                    if (!double.IsNaN(btn.Height) && btn.Height > 0)
                    {
                        btn.MinHeight = Math.Max(58, btn.Height);
                        btn.Height = double.NaN;
                    }
                    if (!double.IsNaN(btn.Width) && btn.Width > 0 && btn.Parent is System.Windows.Controls.Primitives.UniformGrid)
                    {
                        btn.Width = double.NaN;
                    }
                    btn.HorizontalAlignment = HorizontalAlignment.Stretch;
                    btn.Padding = new Thickness(12, 10, 12, 10);
                }
            }
            catch { }
        }

        private void NormalizeSettingsElementRecursive(DependencyObject parent, HashSet<TextBlock> valueBadges)
        {
            if (parent == null) return;

            int count = LogicalTreeHelper.GetChildren(parent).OfType<DependencyObject>().Count();
            foreach (var child in LogicalTreeHelper.GetChildren(parent).OfType<DependencyObject>())
            {
                if (child is TextBlock tb)
                {
                    if (!valueBadges.Contains(tb))
                    {
                        tb.TextWrapping = TextWrapping.Wrap;
                        tb.TextTrimming = TextTrimming.None;
                    }
                    else
                    {
                        tb.TextWrapping = TextWrapping.NoWrap;
                    }
                }
                else if (child is ComboBox cmb)
                {
                    if (double.IsNaN(cmb.MinWidth) || cmb.MinWidth < 225)
                        cmb.MinWidth = 225;
                    if (!double.IsNaN(cmb.Width) && cmb.Width < 225)
                        cmb.Width = 235;
                    if (double.IsNaN(cmb.Height) || cmb.Height < 32)
                        cmb.Height = 34;

                    foreach (var item in cmb.Items.OfType<ComboBoxItem>())
                    {
                        item.Padding = new Thickness(10, 6, 10, 6);
                    }
                }
                else if (child is Border bd)
                {
                    // Sabit yukseklikli ayar satiri kartlarini MinHeight'a cevir ki 2 satira inen aciklamalar kesilmesin
                    if (!double.IsNaN(bd.Height) && bd.Height >= 42 && bd.Height <= 140)
                    {
                        bd.MinHeight = bd.Height;
                        bd.Height = double.NaN;
                    }
                    if (bd.Padding.Top < 8 && bd.CornerRadius.TopLeft >= 10)
                    {
                        bd.Padding = new Thickness(
                            Math.Max(14, bd.Padding.Left),
                            Math.Max(11, bd.Padding.Top),
                            Math.Max(14, bd.Padding.Right),
                            Math.Max(11, bd.Padding.Bottom));
                    }
                }
                else if (child is Grid g && g.ColumnDefinitions.Count >= 2)
                {
                    // Sol aciklama sutununun (*) ve sag kontrol sutununun (Auto) birbirine girmesini engelle
                    var firstColChild = g.Children.OfType<FrameworkElement>().FirstOrDefault(fe => Grid.GetColumn(fe) == 0);
                    if (firstColChild != null && firstColChild.Margin.Right < 14)
                    {
                        firstColChild.Margin = new Thickness(
                            firstColChild.Margin.Left,
                            firstColChild.Margin.Top,
                            18,
                            firstColChild.Margin.Bottom);
                    }
                }

                NormalizeSettingsElementRecursive(child, valueBadges);
            }
        }
        // #endregion SETTINGS_LAYOUT_FIXER

                        // #region SARPNOTE_PRO_FEATURES
        private bool _proFeaturesInitialized = false;
        private bool _isDashedLineMode = false;
        private bool _isVectorArrowMode = false;
        private bool _isSplitViewOpen = false;
        private bool _isQuestionCropMode = false;
        private Point _cropStartPt;
        private InkCanvas _cropTargetInk;
        private RectSelectionAdorner _cropAdorner;
        private Button _btnTopCrop;

        // Deneme Kronometresi (Üst Bar Doğrudan Butonlu)
        private readonly DispatcherTimer _examTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        private bool _examTimerRunning = false;
        private int _examSecondsValue = 0;
        private int _examTargetCountdownSeconds = 0;
        private Button _btnExamToggle;
        private TextBlock _txtExamTimer;

        // Bölünmüş Ekran (Split View - Cevap Anahtarı Paneli)
        private Border _splitViewPanel;
        private Image _splitViewImage;
        private TextBlock _txtSplitPageInfo;
        private ScaleTransform _splitViewScale;
        private int _splitViewPageIndex = 0;
        private double _splitViewZoom = 1.0;

        // Pen Studio Özel Mod Butonları & 3 Favori Kalem Slotu
        private Button _btnToggleDashed;
        private Button _btnToggleArrow;
        private readonly (string Hex, double SizePt)[] _quickPenPresets = new[]
        {
            ("#0F172A", 1.2),
            ("#EF4444", 1.2),
            ("#3B82F6", 1.2)
        };
        private readonly Button[] _presetSlotButtons = new Button[3];

        private void EnsureProFeaturesInitialized()
        {
            if (_proFeaturesInitialized) return;
            _proFeaturesInitialized = true;

            // Sürükle-Bırak (Drag & Drop) PDF Açma
            AllowDrop = true;
            PreviewDragOver += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effects = DragDropEffects.Copy;
                    e.Handled = true;
                }
            };
            PreviewDrop += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] droppedFiles)
                {
                    string lastPdf = null;
                    foreach (string f in droppedFiles)
                    {
                        if (string.Equals(IOPath.GetExtension(f), ".pdf", StringComparison.OrdinalIgnoreCase) && File.Exists(f))
                    {
                        string fn = IOPath.GetFileName(f);
                        string libDest = IOPath.Combine(_libraryFolder, fn);
                        try
                        {
                            if (!string.Equals(f, libDest, StringComparison.OrdinalIgnoreCase))
                                File.Copy(f, libDest, true);
                        }
                        catch { }
                        _meta.OriginalPaths[fn] = File.Exists(libDest) ? libDest : f;
                        lastPdf = fn;
                    }
                    }
                    if (lastPdf != null)
                    {
                        _dashboardDirty = true;
                        SaveMetadata();
                        DashboardOverlay.Visibility = Visibility.Collapsed;
                        OpenPdfFile(lastPdf, 0);
                    }
                    e.Handled = true;
                }
            };

            // Deneme Sayacı Saniye Döngüsü
            _examTimer.Tick += (s, e) =>
            {
                if (!_examTimerRunning) return;
                if (_examTargetCountdownSeconds > 0)
                {
                    if (_examSecondsValue > 0) _examSecondsValue--;
                    else
                    {
                        _examTimerRunning = false;
                        _examTimer.Stop();
                        UpdateExamTimerDisplay();
                        MessageBox.Show("Deneme süresi tamamlandı!", "SarpNote Deneme Sayacı", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                }
                else
                {
                    _examSecondsValue++;
                }
                UpdateExamTimerDisplay();
            };

            BuildTopBarProButtons();
            BuildPenStudioProExtensions();
            BuildSplitViewOverlayPanel();
        }

        private void BuildTopBarProButtons() { }

        private void UpdateExamTimerDisplay()
        {
            if (_txtExamTimer == null) return;
            int hrs = _examSecondsValue / 3600;
            int mins = (_examSecondsValue % 3600) / 60;
            int secs = _examSecondsValue % 60;
            string icon = _examTimerRunning ? "⏸ " : "▶ ";
            _txtExamTimer.Text = hrs > 0
                ? $"{icon}{hrs:00}:{mins:00}:{secs:00}"
                : $"{icon}{mins:00}:{secs:00}";

            _txtExamTimer.Foreground = _examTimerRunning
                ? new SolidColorBrush(Color.FromRgb(110, 231, 183))
                : new SolidColorBrush(Color.FromRgb(226, 232, 240));
        }

        private void BuildPenStudioProExtensions()
        {
            if (ChkHoldToShape?.Parent is not StackPanel studioStack) return;

            var geoModesGrid = new UniformGrid { Columns = 2, Margin = new Thickness(0, 8, 0, 4) };

            _btnToggleDashed = new Button
            {
                Height = 32,
                Margin = new Thickness(0, 0, 4, 0),
                Background = new SolidColorBrush(Color.FromRgb(29, 32, 43)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(46, 52, 72)),
                BorderThickness = new Thickness(1.2),
                Cursor = Cursors.Hand,
                ToolTip = "Kesikli Çizgi Modu (Alt+D)",
                Content = new TextBlock { Text = "- - -  Kesikli Çizgi", Foreground = Brushes.WhiteSmoke, FontSize = 11, FontWeight = FontWeights.SemiBold }
            };
            _btnToggleDashed.Click += (s, e) =>
            {
                _isDashedLineMode = !_isDashedLineMode;
                UpdateGeoModifierButtonsUI();
            };

            _btnToggleArrow = new Button
            {
                Height = 32,
                Margin = new Thickness(4, 0, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(29, 32, 43)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(46, 52, 72)),
                BorderThickness = new Thickness(1.2),
                Cursor = Cursors.Hand,
                ToolTip = "Vektör / Işın Oku Modu (Alt+A)",
                Content = new TextBlock { Text = "➔  Vektör Oku", Foreground = Brushes.WhiteSmoke, FontSize = 11, FontWeight = FontWeights.SemiBold }
            };
            _btnToggleArrow.Click += (s, e) =>
            {
                _isVectorArrowMode = !_isVectorArrowMode;
                UpdateGeoModifierButtonsUI();
            };

            geoModesGrid.Children.Add(_btnToggleDashed);
            geoModesGrid.Children.Add(_btnToggleArrow);
            studioStack.Children.Add(geoModesGrid);

            var presetsHeader = new TextBlock
            {
                Text = "Hızlı Favori Kalemler (Sol Tık: Seç | Sağ Tık: Kaydet)",
                FontSize = 10.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                Margin = new Thickness(2, 6, 0, 4)
            };
            studioStack.Children.Add(presetsHeader);

            var presetGrid = new UniformGrid { Columns = 3, Margin = new Thickness(0, 0, 0, 2) };
            for (int i = 0; i < 3; i++)
            {
                int slotIdx = i;
                var btn = new Button
                {
                    Height = 30,
                    Margin = new Thickness(i == 0 ? 0 : 3, 0, i == 2 ? 0 : 3, 0),
                    Background = new SolidColorBrush(Color.FromRgb(24, 28, 40)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(46, 52, 72)),
                    BorderThickness = new Thickness(1.2),
                    Cursor = Cursors.Hand
                };
                btn.Click += (s, e) => ApplyQuickPenPreset(slotIdx);
                btn.MouseRightButtonUp += (s, e) =>
                {
                    e.Handled = true;
                    string hex = $"#{_penColor.R:X2}{_penColor.G:X2}{_penColor.B:X2}";
                    _quickPenPresets[slotIdx] = (hex, _toolSizes["Pen"]);
                    RefreshQuickPenPresetButtonsUI();
                };
                _presetSlotButtons[i] = btn;
                presetGrid.Children.Add(btn);
            }
            studioStack.Children.Add(presetGrid);
            RefreshQuickPenPresetButtonsUI();
        }

        private void UpdateGeoModifierButtonsUI()
        {
            var actBg = new SolidColorBrush(Color.FromRgb(37, 42, 61));
            var normBg = new SolidColorBrush(Color.FromRgb(29, 32, 43));
            var actBd = new SolidColorBrush(Color.FromRgb(99, 102, 241));
            var normBd = new SolidColorBrush(Color.FromRgb(46, 52, 72));

            if (_btnToggleDashed != null)
            {
                _btnToggleDashed.Background = _isDashedLineMode ? actBg : normBg;
                _btnToggleDashed.BorderBrush = _isDashedLineMode ? actBd : normBd;
            }
            if (_btnToggleArrow != null)
            {
                _btnToggleArrow.Background = _isVectorArrowMode ? actBg : normBg;
                _btnToggleArrow.BorderBrush = _isVectorArrowMode ? actBd : normBd;
            }
        }

        private void RefreshQuickPenPresetButtonsUI()
        {
            for (int i = 0; i < 3; i++)
            {
                var btn = _presetSlotButtons[i];
                if (btn == null) continue;
                var preset = _quickPenPresets[i];
                Color c = (Color)ColorConverter.ConvertFromString(preset.Hex);
                var sp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                sp.Children.Add(new System.Windows.Shapes.Ellipse
                {
                    Width = 9, Height = 9,
                    Fill = new SolidColorBrush(c),
                    Margin = new Thickness(0, 0, 6, 0)
                });
                sp.Children.Add(new TextBlock
                {
                    Text = $"{preset.SizePt:0.#} pt",
                    Foreground = Brushes.WhiteSmoke,
                    FontSize = 11,
                    FontWeight = FontWeights.Bold
                });
                btn.Content = sp;
            }
        }

        private void ApplyQuickPenPreset(int index)
        {
            if (index < 0 || index >= _quickPenPresets.Length) return;
            var preset = _quickPenPresets[index];
            _isQuestionCropMode = false;
            UpdateCropButtonVisualState();
            _activeTool = "Pen";
            _activeShapePreset = 0;
            _penColor = _activeColor = (Color)ColorConverter.ConvertFromString(preset.Hex);
            _penThickness = _toolSizes["Pen"] = preset.SizePt;
            HighlightActiveToolButtons();
            UpdatePenStudioVisuals();
            ApplyToolSettingsToAllRealizedPages();
        }

        private void ApplySpecialLineModifiersIfEnabled(InkCanvas ink, Stroke stroke)
        {
            if (ink == null || stroke == null || _activeTool == "Highlighter") return;
            if (!_isDashedLineMode && !_isVectorArrowMode) return;

            var verts = GetStrokeVertices(stroke);
            string kind = stroke.ContainsPropertyData(GeoShapeTypeGuid)
                ? stroke.GetPropertyData(GeoShapeTypeGuid) as string
                : null;

            if (verts == null || verts.Count < 2) return;

            _suppressHistory = true;
            try
            {
                if (_isVectorArrowMode && (kind == "Line" || kind == "Polyline") && !_isDashedLineMode)
                {
                    Point pStart = verts[verts.Count - 2];
                    Point pEnd = verts[verts.Count - 1];
                    if (Dist(pStart, pEnd) >= 14.0)
                    {
                        var arrowPts = new StylusPointCollection(stroke.StylusPoints);
                        AppendArrowheadPoints(arrowPts, pStart, pEnd);
                        stroke.StylusPoints = arrowPts;
                    }
                }

                if (_isDashedLineMode && (kind == "Line" || kind == "Polyline" || kind == "Triangle" || kind == "Rectangle" || kind == "Square"))
                {
                    var poly = new List<Point>(verts);
                    if (kind == "Triangle" || kind == "Rectangle" || kind == "Square")
                        poly.Add(verts[0]);

                    double dashLen = Math.Max(7.0, _penThickness * 4.2);
                    double gapLen = Math.Max(5.5, _penThickness * 3.2);
                    var da = stroke.DrawingAttributes.Clone();
                    da.FitToCurve = false;

                    var dashedStrokes = new StrokeCollection();
                    for (int s = 0; s < poly.Count - 1; s++)
                    {
                        Point a = poly[s];
                        Point b = poly[s + 1];
                        double total = Dist(a, b);
                        if (total < 1.0) continue;

                        double ux = (b.X - a.X) / total;
                        double uy = (b.Y - a.Y) / total;
                        double pos = 0.0;

                        while (pos < total)
                        {
                            double segEnd = Math.Min(total, pos + dashLen);
                            if (segEnd - pos >= 1.5)
                            {
                                Point d0 = new Point(a.X + ux * pos, a.Y + uy * pos);
                                Point d1 = new Point(a.X + ux * segEnd, a.Y + uy * segEnd);
                                dashedStrokes.Add(new Stroke(CreateDenseLineStylusPoints(d0, d1), da.Clone()));
                            }
                            pos += dashLen + gapLen;
                        }
                    }

                    if (_isVectorArrowMode && poly.Count >= 2)
                    {
                        Point pStart = poly[poly.Count - 2];
                        Point pEnd = poly[poly.Count - 1];
                        var headPts = new StylusPointCollection { new StylusPoint(pEnd.X, pEnd.Y) };
                        AppendArrowheadPoints(headPts, pStart, pEnd);
                        dashedStrokes.Add(new Stroke(headPts, da.Clone()));
                    }

                    if (dashedStrokes.Count > 0)
                    {
                        ink.Strokes.Remove(stroke);
                        ink.Strokes.Add(dashedStrokes);
                    }
                }
            }
            finally
            {
                _suppressHistory = false;
            }
        }

        private void AppendArrowheadPoints(StylusPointCollection pts, Point pStart, Point pEnd)
        {
            double angle = Math.Atan2(pEnd.Y - pStart.Y, pEnd.X - pStart.X);
            double headLen = Math.Clamp(10.0 + _penThickness * 2.2, 11.0, 24.0);
            double wingAngle = 27.0 * (Math.PI / 180.0);

            Point wing1 = new Point(
                pEnd.X - headLen * Math.Cos(angle - wingAngle),
                pEnd.Y - headLen * Math.Sin(angle - wingAngle));
            Point wing2 = new Point(
                pEnd.X - headLen * Math.Cos(angle + wingAngle),
                pEnd.Y - headLen * Math.Sin(angle + wingAngle));

            foreach (var sp in CreateDenseLineStylusPoints(pEnd, wing1)) pts.Add(sp);
            foreach (var sp in CreateDenseLineStylusPoints(wing1, pEnd)) pts.Add(sp);
            foreach (var sp in CreateDenseLineStylusPoints(pEnd, wing2)) pts.Add(sp);
        }

        // Bölünmüş Ekran (Split View - Cevap Anahtarı Paneli)
        private void BuildSplitViewOverlayPanel()
        {
            if (WorkspaceGrid == null || _splitViewPanel != null) return;

            _splitViewPanel = new Border
            {
                Width = 440,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Stretch,
                Margin = new Thickness(0, 8, 68, 52),
                Background = new SolidColorBrush(Color.FromRgb(18, 21, 32)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(63, 72, 104)),
                BorderThickness = new Thickness(1.5),
                CornerRadius = new CornerRadius(14),
                Visibility = Visibility.Collapsed,
                ClipToBounds = true
            };
            _splitViewPanel.Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                Color = Colors.Black,
                BlurRadius = 28,
                ShadowDepth = 6,
                Opacity = 0.75
            };

            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(38) });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var header = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(24, 28, 43)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(42, 48, 70)),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10, 0, 8, 0)
            };

            var hGrid = new Grid();
            hGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            hGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var title = new TextBlock
            {
                Text = "◫ Cevap Anahtarı / Bölünmüş Sayfa",
                Foreground = Brushes.WhiteSmoke,
                FontSize = 11.5,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(title, 0);

            var navStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            var ghostStyle = TryFindResource("TopGhostBtn") as Style;

            var btnPrev = new Button { Content = "◀", Style = ghostStyle, Width = 24, Height = 24, ToolTip = "Önceki Sayfa" };
            btnPrev.Click += (s, e) => ChangeSplitViewPage(_splitViewPageIndex - 1);

            _txtSplitPageInfo = new TextBlock
            {
                Text = "1 / 1",
                Foreground = new SolidColorBrush(Color.FromRgb(165, 180, 252)),
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 6, 0)
            };

            var btnNext = new Button { Content = "▶", Style = ghostStyle, Width = 24, Height = 24, ToolTip = "Sonraki Sayfa" };
            btnNext.Click += (s, e) => ChangeSplitViewPage(_splitViewPageIndex + 1);

            var btnLast = new Button { Content = "Son", Style = ghostStyle, Height = 24, Padding = new Thickness(5, 0, 5, 0), Margin = new Thickness(3, 0, 3, 0), ToolTip = "En Son Sayfaya (Cevap Anahtarına) Git" };
            btnLast.Click += (s, e) =>
            {
                if (_currentPdf != null) ChangeSplitViewPage(_currentPdf.PageCount - 1);
            };

            var btnClose = new Button { Content = "✕", Style = ghostStyle, Width = 24, Height = 24, ToolTip = "Bölünmüş Ekranı Kapat" };
            btnClose.Click += (s, e) => ToggleSplitViewPanel();

            navStack.Children.Add(btnPrev);
            navStack.Children.Add(_txtSplitPageInfo);
            navStack.Children.Add(btnNext);
            navStack.Children.Add(btnLast);
            navStack.Children.Add(btnClose);
            Grid.SetColumn(navStack, 1);

            hGrid.Children.Add(title);
            hGrid.Children.Add(navStack);
            header.Child = hGrid;
            Grid.SetRow(header, 0);

            var sv = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = new SolidColorBrush(Color.FromRgb(12, 14, 22))
            };
            _splitViewScale = new ScaleTransform(1.0, 1.0);
            _splitViewImage = new Image
            {
                Stretch = Stretch.Uniform,
                LayoutTransform = _splitViewScale
            };
            RenderOptions.SetBitmapScalingMode(_splitViewImage, BitmapScalingMode.HighQuality);
            sv.PreviewMouseWheel += (s, e) =>
            {
                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    _splitViewZoom = Math.Clamp(_splitViewZoom * (e.Delta > 0 ? 1.12 : 0.89), 0.6, 3.0);
                    _splitViewScale.ScaleX = _splitViewScale.ScaleY = _splitViewZoom;
                    e.Handled = true;
                }
            };
            sv.Content = _splitViewImage;
            Grid.SetRow(sv, 1);

            root.Children.Add(header);
            root.Children.Add(sv);
            _splitViewPanel.Child = root;
            WorkspaceGrid.Children.Add(_splitViewPanel);
        }

        private void ToggleSplitViewPanel()
        {
            if (_splitViewPanel == null) BuildSplitViewOverlayPanel();
            if (_splitViewPanel == null || _currentPdf == null) return;

            _isSplitViewOpen = !_isSplitViewOpen;
            _splitViewPanel.Visibility = _isSplitViewOpen ? Visibility.Visible : Visibility.Collapsed;
            if (_isSplitViewOpen)
            {
                int target = _splitViewPageIndex > 0 ? _splitViewPageIndex : Math.Max(0, _currentPdf.PageCount - 1);
                ChangeSplitViewPage(target);
            }
        }

        private void ChangeSplitViewPage(int pageIdx)
        {
            var session = _activeSession;
            if (session?.Document == null || _splitViewImage == null) return;

            _splitViewPageIndex = Math.Clamp(pageIdx, 0, session.Document.PageCount - 1);
            if (_txtSplitPageInfo != null)
                _txtSplitPageInfo.Text = $"{_splitViewPageIndex + 1} / {session.Document.PageCount}";

            if (session.HighResCache.TryGetValue(_splitViewPageIndex, out var cached))
            {
                _splitViewImage.Source = cached;
                return;
            }

            int p = _splitViewPageIndex;
            Task.Run(() =>
            {
                try
                {
                    BitmapSource bmpSrc;
                    lock (session.SyncLock)
                    {
                        if (session.Document == null) return;
                        var sz = session.Document.PageSizes[p];
                        int w = (int)(GetNormalizedPageWidth(sz) * 1.6);
                        int h = (int)(GetNormalizedPageHeight(sz) * 1.6);
                        using var bmp = (GdiBitmap)session.Document.RenderCrisp(p, w, h, 144, 144, PdfRenderFlags.None);
                        bmpSrc = FastBitmapToSource(bmp);
                    }
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (_splitViewImage != null && _splitViewPageIndex == p)
                            _splitViewImage.Source = bmpSrc;
                    }));
                }
                catch { }
            });
        }

        // =====================================================================
        // SORU KIRPMA MOTORU (STYLUS VE MOUSE TAM UYUMLU + ONIZLEME PENCERESI)
        // =====================================================================
        private void UpdateCropButtonVisualState()
        {
            if (_btnTopCrop == null) return;
            _btnTopCrop.Background = _isQuestionCropMode
                ? new SolidColorBrush(Color.FromRgb(30, 58, 138))
                : Brushes.Transparent;
        }

        private void StartQuestionCropMode() => ToggleQuestionCropMode(forceState: true);

        private void ToggleQuestionCropMode(bool? forceState = null)
        {
            if (_currentPdf == null) return;
            _isQuestionCropMode = forceState ?? !_isQuestionCropMode;
            UpdateCropButtonVisualState();

            if (_isQuestionCropMode)
            {
                if (WorkspaceGrid != null) WorkspaceGrid.Cursor = Cursors.Cross;
                foreach (var s in _pageSlots)
                {
                    if (s.IsRealized && s.Canvas != null)
                        s.Canvas.Cursor = Cursors.Cross;
                }
            }
            else
            {
                _cropTargetInk = null;
                ApplyToolSettingsToAllRealizedPages();
            }
        }

        private void BeginCropOnCanvas(InkCanvas ink, Point pt, bool isStylus)
        {
            if (ink == null) return;
            _cropTargetInk = ink;
            _cropStartPt = pt;

            var layer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(ink);
            if (layer != null)
            {
                if (_cropAdorner != null) { try { layer.Remove(_cropAdorner); } catch { } }
                _cropAdorner = new RectSelectionAdorner(ink, _cropStartPt);
                layer.Add(_cropAdorner);
            }

            if (isStylus) ink.CaptureStylus();
            else ink.CaptureMouse();
        }

        private void UpdateCropOnCanvas(Point pt)
        {
            _cropAdorner?.UpdatePoint(pt);
        }

                private void CompleteCropOnCanvas(Point endPt)
        {
            var ink = _cropTargetInk;
            if (ink == null) return;

            try
            {
                if (_cropAdorner != null)
                {
                    var layer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(ink);
                    if (layer != null) { try { layer.Remove(_cropAdorner); } catch { } }
                    _cropAdorner = null;
                }
                Stylus.Captured?.ReleaseStylusCapture();
                if (ink.IsMouseCaptured) ink.ReleaseMouseCapture();

                double x1 = Math.Min(_cropStartPt.X, endPt.X);
                double y1 = Math.Min(_cropStartPt.Y, endPt.Y);
                double x2 = Math.Max(_cropStartPt.X, endPt.X);
                double y2 = Math.Max(_cropStartPt.Y, endPt.Y);

                if ((x2 - x1) >= 14 && (y2 - y1) >= 14 && ink.Tag is PageSlot slot && slot.Container != null)
                {
                    if (!slot.IsRealized) RealizePageSlot(slot);

                    double dpiScale = 2.0;
                    var fullPageBmp = RenderSlotToBitmap(slot, dpiScale);

                    int pxX = Math.Clamp((int)Math.Round(x1 * dpiScale), 0, Math.Max(0, fullPageBmp.PixelWidth - 2));
                    int pxY = Math.Clamp((int)Math.Round(y1 * dpiScale), 0, Math.Max(0, fullPageBmp.PixelHeight - 2));
                    int pxW = Math.Clamp((int)Math.Round((x2 - x1) * dpiScale), 1, fullPageBmp.PixelWidth - pxX);
                    int pxH = Math.Clamp((int)Math.Round((y2 - y1) * dpiScale), 1, fullPageBmp.PixelHeight - pxY);

                    var cropped = new CroppedBitmap(fullPageBmp, new Int32Rect(pxX, pxY, pxW, pxH));
                    cropped.Freeze();

                    // Panoya saf RGB/BGR formatında kopyala ki yapıştırılan uygulamalarda (WhatsApp/Paint/Word) siyah arka plan çıkmasın
                    var opaqueBgr = new FormatConvertedBitmap(cropped, PixelFormats.Bgr32, null, 0);
                    opaqueBgr.Freeze();
                    Clipboard.SetImage(opaqueBgr);

                    ShowCroppedQuestionDialog(opaqueBgr, slot.PageIndex + 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kırpma sırasında hata oluştu: {ex.Message}", "SarpNote", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                _isQuestionCropMode = false;
                _cropTargetInk = null;
                UpdateCropButtonVisualState();
                ApplyToolSettingsToAllRealizedPages();
            }
        }

        private void ShowCroppedQuestionDialog(BitmapSource croppedImg, int pageNumber)
        {
            var dlg = new Window
            {
                Title = TrText("SarpNote - Kırpılan Soru"),
                Width = 520,
                Height = 440,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Background = new SolidColorBrush(Color.FromRgb(15, 18, 28)),
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var infoTxt = new TextBlock
            {
                Text = TrText("✓ Soru panoya kopyalandı (Ctrl+V ile yapıştırabilir veya aşağıdan kaydedebilirsin):"),
                Foreground = new SolidColorBrush(Color.FromRgb(110, 231, 183)),
                FontSize = 12.5,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(infoTxt, 0);

            var imgBorder = new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(6),
                Child = new Image { Source = croppedImg, Stretch = Stretch.Uniform }
            };
            Grid.SetRow(imgBorder, 1);

            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };

            var btnSaveHata = new Button
            {
                Content = TrText("📌 Hata Defterine Kaydet"),
                Height = 34,
                Padding = new Thickness(14, 0, 14, 0),
                Margin = new Thickness(0, 0, 8, 0),
                Background = new SolidColorBrush(Color.FromRgb(37, 99, 235)),
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                Cursor = Cursors.Hand
            };
            btnSaveHata.Click += (s, e) =>
            {
                string tag = (!string.IsNullOrEmpty(_currentPdfFileName) && _meta.Tags.TryGetValue(_currentPdfFileName, out string t)) ? t : "Genel";
                string dir = IOPath.Combine(_baseDir, "HataDefteri", tag);
                Directory.CreateDirectory(dir);
                string filePath = IOPath.Combine(dir, $"{_currentPdfKey}_Sayfa{pageNumber}_{DateTime.Now:HHmmss}.png");
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(croppedImg));
                using (var fs = new FileStream(filePath, FileMode.Create)) enc.Save(fs);
                dlg.Close();
                try { Process.Start("explorer.exe", $"/select,\"{filePath}\""); } catch { }
            };

            var btnSaveAs = new Button
            {
                Content = TrText("💾 Farklı Kaydet (PNG)..."),
                Height = 34,
                Padding = new Thickness(14, 0, 14, 0),
                Margin = new Thickness(0, 0, 8, 0),
                Background = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                Cursor = Cursors.Hand
            };
            btnSaveAs.Click += (s, e) =>
            {
                var sfd = new SaveFileDialog
                {
                    Filter = "PNG Görseli (*.png)|*.png",
                    FileName = $"{_currentPdfKey}_Sayfa{pageNumber}_Soru.png"
                };
                if (sfd.ShowDialog() == true)
                {
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(croppedImg));
                    using (var fs = new FileStream(sfd.FileName, FileMode.Create)) enc.Save(fs);
                    dlg.Close();
                }
            };

            var btnClose = new Button
            {
                Content = TrText("Kapat"),
                Height = 34,
                Padding = new Thickness(14, 0, 14, 0),
                Background = new SolidColorBrush(Color.FromRgb(51, 65, 85)),
                Foreground = Brushes.White,
                Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => dlg.Close();

            btnPanel.Children.Add(btnSaveHata);
            btnPanel.Children.Add(btnSaveAs);
            btnPanel.Children.Add(btnClose);
            Grid.SetRow(btnPanel, 2);

            root.Children.Add(infoTxt);
            root.Children.Add(imgBorder);
            root.Children.Add(btnPanel);
            dlg.Content = root;
            dlg.ShowDialog();
        }

        // =====================================================================
        // COZUMLU PDF VE PNG DISA AKTARMA MOTORU (FARKLI KAYDET DIALOGLU)
        // =====================================================================
        private void ExportAnnotatedPages()
        {
            if (_activeSession == null || _pageSlots.Count == 0) return;
            SaveAllRealizedStrokes();

            var sfd = new SaveFileDialog
            {
                Title = "Çözümlü PDF veya Sayfa Görseli Olarak Dışa Aktar",
                Filter = "Çözümlü PDF - Sadece İşlem Yapılan Sayfalar (*.pdf)|*.pdf|Çözümlü PDF - Tüm Kitap/Deneme (*.pdf)|*.pdf|Mevcut Sayfa Yüksek Çözünürlüklü PNG (*.png)|*.png",
                FilterIndex = 1,
                FileName = $"{_currentPdfKey}_Cozumlu"
            };

            if (sfd.ShowDialog() != true) return;

            try
            {
                if (sfd.FilterIndex == 3 || sfd.FileName.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                {
                    var curSlot = GetActiveSlot() ?? _pageSlots[Math.Clamp(_currentPage, 0, _pageSlots.Count - 1)];
                    bool wasRealized = curSlot.IsRealized;
                    if (!wasRealized) RealizePageSlot(curSlot);

                    var rtb = RenderSlotToBitmap(curSlot, 2.0);
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(rtb));
                    using (var fs = new FileStream(sfd.FileName, FileMode.Create))
                        enc.Save(fs);

                    if (!wasRealized) UnrealizePageSlot(curSlot);
                    try { Process.Start(new ProcessStartInfo { FileName = sfd.FileName, UseShellExecute = true }); } catch { }
                    return;
                }

                bool exportAllPages = (sfd.FilterIndex == 2);
                var pdfPagesData = new List<(byte[] JpegBytes, int PixelW, int PixelH, double PtW, double PtH)>();

                foreach (var slot in _pageSlots)
                {
                    bool hasStrokes = (slot.IsRealized && slot.Canvas != null && slot.Canvas.Strokes.Count > 0)
                        || (slot.CachedStrokeBytes != null && slot.CachedStrokeBytes.Length > 0);

                    if (!exportAllPages && !hasStrokes && slot.PageIndex != _currentPage)
                        continue;

                    bool wasRealized = slot.IsRealized;
                    if (!wasRealized) RealizePageSlot(slot);

                    if (slot.PdfImage != null && slot.PdfImage.Source == null && _activeSession.Document != null)
                    {
                        lock (_activeSession.SyncLock)
                        {
                            using var bmp = (GdiBitmap)_activeSession.Document.RenderCrisp(slot.PageIndex, (int)(slot.PdfWidth * 1.8), (int)(slot.PdfHeight * 1.8), 144, 144, PdfRenderFlags.None);
                            slot.PdfImage.Source = FastBitmapToSource(bmp);
                        }
                    }

                    var rtb = RenderSlotToBitmap(slot, 1.75);
                    var jpgEnc = new JpegBitmapEncoder { QualityLevel = 92 };
                    jpgEnc.Frames.Add(BitmapFrame.Create(rtb));
                    using var ms = new MemoryStream();
                    jpgEnc.Save(ms);

                    pdfPagesData.Add((ms.ToArray(), rtb.PixelWidth, rtb.PixelHeight, slot.TotalWidth * 0.75, slot.PdfHeight * 0.75));

                    if (!wasRealized) UnrealizePageSlot(slot);
                }

                if (pdfPagesData.Count > 0)
                {
                    WriteImagesAsPdfDocument(sfd.FileName, pdfPagesData);
                    try { Process.Start(new ProcessStartInfo { FileName = sfd.FileName, UseShellExecute = true }); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Dışa aktarma sırasında hata oluştu: {ex.Message}", "SarpNote", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private RenderTargetBitmap RenderSlotToBitmap(PageSlot slot, double dpiScale)
        {
            slot.Container.Measure(new Size(slot.TotalWidth, slot.PdfHeight));
            slot.Container.Arrange(new Rect(0, 0, slot.TotalWidth, slot.PdfHeight));
            slot.Container.UpdateLayout();

            int pxW = Math.Max(1, (int)Math.Round(slot.TotalWidth * dpiScale));
            int pxH = Math.Max(1, (int)Math.Round(slot.PdfHeight * dpiScale));

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, slot.TotalWidth, slot.PdfHeight));
                var vb = new VisualBrush(slot.Container) { Stretch = Stretch.None };
                dc.DrawRectangle(vb, null, new Rect(0, 0, slot.TotalWidth, slot.PdfHeight));
            }

            var rtb = new RenderTargetBitmap(pxW, pxH, 96.0 * dpiScale, 96.0 * dpiScale, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        private static void WriteImagesAsPdfDocument(
            string outputPdfPath,
            List<(byte[] JpegBytes, int PixelW, int PixelH, double PtW, double PtH)> pages)
        {
            using var fs = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write);
            void WriteAscii(string s)
            {
                byte[] b = System.Text.Encoding.ASCII.GetBytes(s);
                fs.Write(b, 0, b.Length);
            }

            int pageCount = pages.Count;
            int totalObjects = 3 + pageCount * 3;
            long[] offsets = new long[totalObjects];

            WriteAscii("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");

            // 1: Catalog
            offsets[1] = fs.Position;
            WriteAscii("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

            // 2: Pages
            offsets[2] = fs.Position;
            var kids = string.Join(" ", Enumerable.Range(0, pageCount).Select(i => $"{3 + i * 3} 0 R"));
            WriteAscii($"2 0 obj\n<< /Type /Pages /Kids [ {kids} ] /Count {pageCount} >>\nendobj\n");

            for (int i = 0; i < pageCount; i++)
            {
                var p = pages[i];
                int pageObjId = 3 + i * 3;
                int contentObjId = 4 + i * 3;
                int imgObjId = 5 + i * 3;

                string ptWStr = p.PtW.ToString("0.##", CultureInfo.InvariantCulture);
                string ptHStr = p.PtH.ToString("0.##", CultureInfo.InvariantCulture);

                offsets[pageObjId] = fs.Position;
                WriteAscii($"{pageObjId} 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im0 {imgObjId} 0 R >> >> /MediaBox [0 0 {ptWStr} {ptHStr}] /Contents {contentObjId} 0 R >>\nendobj\n");

                string contentCmd = $"q {ptWStr} 0 0 {ptHStr} 0 0 cm /Im0 Do Q";
                offsets[contentObjId] = fs.Position;
                WriteAscii($"{contentObjId} 0 obj\n<< /Length {contentCmd.Length} >>\nstream\n{contentCmd}\nendstream\nendobj\n");

                offsets[imgObjId] = fs.Position;
                WriteAscii($"{imgObjId} 0 obj\n<< /Type /XObject /Subtype /Image /Width {p.PixelW} /Height {p.PixelH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {p.JpegBytes.Length} >>\nstream\n");
                fs.Write(p.JpegBytes, 0, p.JpegBytes.Length);
                WriteAscii("\nendstream\nendobj\n");
            }

            long xrefPos = fs.Position;
            WriteAscii($"xref\n0 {totalObjects}\n0000000000 65535 f \n");
            for (int i = 1; i < totalObjects; i++)
            {
                WriteAscii($"{offsets[i]:0000000000} 00000 n \n");
            }
            WriteAscii($"trailer\n<< /Size {totalObjects} /Root 1 0 R >>\nstartxref\n{xrefPos}\n%%EOF\n");
        }
        // #endregion SARPNOTE_PRO_FEATURES

                                                // #region SARPNOTE_LOCALIZATION_ENGINE
        private string _currentLang = "tr";
        private bool _isApplyingLanguage = false;
        private bool _localizationHooksAttached = false;
        private ComboBox _cmbAppLanguage;
        private Border _langSettingCard;

        private readonly Dictionary<DependencyObject, string> _originalTextMap = new Dictionary<DependencyObject, string>();
        private readonly Dictionary<DependencyObject, string> _lastAppliedTextMap = new Dictionary<DependencyObject, string>();
        private readonly Dictionary<DependencyObject, string> _originalToolTipMap = new Dictionary<DependencyObject, string>();
        private readonly Dictionary<DependencyObject, string> _lastAppliedToolTipMap = new Dictionary<DependencyObject, string>();
        private readonly HashSet<DependencyObject> _structurallyTranslatedElements = new HashSet<DependencyObject>();

        private static readonly Dictionary<string, Dictionary<string, string>> _i18nDict = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["tr"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PAGES"] = "SAYFALAR",
                ["Pages"] = "Sayfalar",
                ["PAGE"] = "SAYFA",
                ["Pen Studio"] = "Kalem Stüdyosu",
                ["Basic"] = "Basit",
                ["Pro"] = "Pro"
            },

            ["en"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PAGES"] = "PAGES",
                ["SAYFALAR"] = "PAGES",
                ["Sayfalar"] = "Pages",
                ["Sayfa"] = "Page",
                ["Tüm Kitaplar"] = "All Books",
                ["Yıldızlı Sorularım"] = "Starred Questions",
                ["Matematik & Geometri"] = "Mathematics & Geometry",
                ["TYT / AYT Deneme"] = "TYT / AYT Practice Exam",
                ["Kitap adı, klasör veya ders ara..."] = "Search book name, folder or subject...",
                ["+ Kitap Ekle"] = "+ Add Book",
                ["Kitap Ekle"] = "Add Book",
                ["Çizime Dön"] = "Back to Drawing",
                ["Kitap Adı"] = "Book Name",
                ["Klasör Konumu"] = "Folder Location",
                ["Kaldığın Sayfa"] = "Last Page",
                ["Son Açılma"] = "Last Opened",
                ["Ders"] = "Subject",
                ["Dersler"] = "Subjects",
                ["Kategoriler"] = "Categories",
                ["Bugün"] = "Today",
                ["Dün"] = "Yesterday",
                ["Az önce"] = "Just now",
                ["Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran"] = "Auto Solution Sheet & Fullscreen on New PDFs",
                ["Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar."] = "Automatically prepares the right-side scratchpad for every opened test.",
                ["Kapalı"] = "Off",
                ["Açık"] = "On",
                ["Bölünmüş"] = "Split View",
                ["◫ Bölünmüş"] = "◫ Split View",
                ["Kırp"] = "Crop",
                ["✂ Kırp"] = "✂ Crop",
                ["Dışa Aktar"] = "Export",
                ["📤 Dışa Aktar"] = "📤 Export",
                ["Kitaplık"] = "Library",
                ["KÜTÜPHANE"] = "LIBRARY",
                ["KİTAPLIK"] = "LIBRARY",
                ["KLASÖRLER"] = "FOLDERS",
                ["Klasörler"] = "Folders",
                ["ETİKETLER"] = "TAGS",
                ["DERS ETİKETLERİ"] = "SUBJECT TAGS",
                ["Ders Etiketleri"] = "Subject Tags",
                ["Etiketler"] = "Tags",
                ["Tüm Belgeler"] = "All Documents",
                ["Tüm PDF'ler"] = "All PDFs",
                ["Son Açılanlar"] = "Recently Opened",
                ["Son Okunanlar"] = "Recently Read",
                ["Favoriler"] = "Favorites",
                ["Yıldızlılar"] = "Starred",
                ["Yıldızlı Sayfalar"] = "Starred Pages",
                ["PDF Ekle"] = "Add PDF",
                ["+ PDF Ekle"] = "+ Add PDF",
                ["PDF Yükle"] = "Import PDF",
                ["+ PDF Yükle"] = "+ Import PDF",
                ["Yeni Klasör"] = "New Folder",
                ["+ Yeni Klasör"] = "+ New Folder",
                ["+ Klasör"] = "+ Folder",
                ["Klasör Oluştur"] = "Create Folder",
                ["Yeni Not Defteri"] = "New Notebook",
                ["Kitaplıkta ara..."] = "Search in library...",
                ["Belge veya etiket ara..."] = "Search documents or tags...",
                ["Ara..."] = "Search...",
                ["Henüz PDF eklenmedi"] = "No PDFs added yet",
                ["Henüz belge bulunmuyor"] = "No documents found",
                ["Bu klasörde belge yok"] = "No documents in this folder",
                ["Aç"] = "Open",
                ["Sil"] = "Delete",
                ["Kaldır"] = "Remove",
                ["Listeden Kaldır"] = "Remove from List",
                ["Favorilere Ekle"] = "Add to Favorites",
                ["Favorilerden Çıkar"] = "Remove from Favorites",
                ["Klasöre Taşı"] = "Move to Folder",
                ["Etiket Değiştir"] = "Change Tag",
                ["Dosya Konumunu Aç"] = "Open File Location",
                ["Genel"] = "General",
                ["Matematik"] = "Mathematics",
                ["Geometri"] = "Geometry",
                ["Fizik"] = "Physics",
                ["Kimya"] = "Chemistry",
                ["Biyoloji"] = "Biology",
                ["Türkçe"] = "Turkish",
                ["Tarih"] = "History",
                ["Coğrafya"] = "Geography",
                ["Felsefe"] = "Philosophy",
                ["Deneme"] = "Practice Exam",
                ["Tümü"] = "All",
                ["AYARLAR"] = "SETTINGS",
                ["Ayarlar"] = "Settings",
                ["Uygulama Ayarları"] = "Application Settings",
                ["SarpNote Ayarları"] = "SarpNote Settings",
                ["Kaydet ve Kapat"] = "Save & Close",
                ["Tamam"] = "OK",
                ["Uygula"] = "Apply",
                ["İptal"] = "Cancel",
                ["Kapat"] = "Close",
                ["Varsayılana Sıfırla"] = "Reset to Defaults",
                ["Varsayılanlara Sıfırla"] = "Reset to Defaults",
                ["Ayarları Sıfırla"] = "Reset Settings",
                ["Sıfırla"] = "Reset",
                ["Kalem & Silgi"] = "Pen & Eraser",
                ["Kalem ve Silgi"] = "Pen & Eraser",
                ["Kalem & Silgi Motoru"] = "Pen & Eraser Engine",
                ["Kalem ve Silgi Motoru"] = "Pen & Eraser Engine",
                ["Çizim hassasiyeti, şekil tanıma ve silgi davranışlarını özelleştirin."] = "Customize drawing precision, shape recognition, and eraser behavior.",
                ["Görüntü & Performans"] = "Display & Performance",
                ["Görüntü ve Performans"] = "Display & Performance",
                ["4K PDF render kalitesi, okuma filtreleri ve kaydırma akıcılığını ayarlayın."] = "Adjust 4K PDF render quality, reading filters, and scrolling smoothness.",
                ["Çalışma Alanı"] = "Workspace",
                ["Çalışma Alanı & Kağıt"] = "Workspace & Paper",
                ["Kağıt & Çalışma Alanı"] = "Paper & Workspace",
                ["Varsayılan yakınlaştırma, ek çözüm kağıdı ve kılavuz çizgi tercihlerini yönetin."] = "Manage default zoom, extra solution sheet, and guide line preferences.",
                ["Depolama"] = "Storage",
                ["Depolama & Veri"] = "Storage & Data",
                ["Depolama ve Veri Yönetimi"] = "Storage & Data Management",
                ["Kütüphane istatistiklerini görüntüleyin, önbelleği ve kayıt klasörünü yönetin."] = "View library statistics, manage cache and storage directory.",
                ["Başlangıç"] = "Startup",
                ["Başlangıç Ayarları"] = "Startup Settings",
                ["Uygulama açılışındaki varsayılan ekran, araç ve pencere tercihlerini belirleyin."] = "Set default screen, tool, language, and window preferences on startup.",
                ["Varsayılan Kalem Modu"] = "Default Pen Mode",
                ["Çizim yaparken kullanılacak varsayılan kalem motorunu seçin."] = "Select the default pen engine used while drawing.",
                ["Basic"] = "Basic",
                ["Pro"] = "Pro",
                ["Basit"] = "Basic",
                ["Basic (Sabit Uç)"] = "Basic (Fixed Tip)",
                ["Basit (Sabit Kalınlık)"] = "Basic (Fixed Width)",
                ["Basınçsız, sabit kalınlık"] = "No pressure, uniform width",
                ["Pro (Basınç Duyarlı)"] = "Pro (Pressure Sensitive)",
                ["Basınç destekli dinamik uç"] = "Pressure-sensitive dynamic tip",
                ["Basılı Tutunca Şekle Dönüştür"] = "Hold to Convert to Shape",
                ["Basılı tutunca otomatik şekil"] = "Hold for auto-shape",
                ["Çizgi veya kapalı şekil çizip ucunda bekletince düzgün geometriye dönüştürür."] = "Converts lines or closed shapes into crisp geometry when held at the end.",
                ["Basılı Tutma Süresi (Otomatik Şekil)"] = "Hold Duration (Auto-Shape)",
                ["Basılı Tutma Süresi"] = "Hold Duration",
                ["Çizgiyi düzleştirmek veya şekle dönüştürmek için gereken bekleme süresi."] = "Wait time required to straighten a line or convert a drawn shape.",
                ["Çizgi Yumuşatma (Smoothing)"] = "Stroke Smoothing",
                ["Çizgi Yumuşatma"] = "Stroke Smoothing",
                ["El titremesini azaltır ve çizgilerin daha pürüzsüz görünmesini sağlar."] = "Reduces hand jitter and makes strokes look smoother.",
                ["Varsayılan Silgi Modu"] = "Default Eraser Mode",
                ["Silgi Modu"] = "Eraser Mode",
                ["Silginin çizgileri tek dokunuşla mı yoksa bölgesel mi sileceğini belirler."] = "Determines whether the eraser removes whole strokes or partial points.",
                ["Çizgi Silgisi (Tüm Çizgiyi Siler)"] = "Stroke Eraser (Erases Entire Stroke)",
                ["Nokta Silgisi (Kısmi Siler)"] = "Point Eraser (Partial Erase)",
                ["Nokta Silgisi (Hassas)"] = "Point Eraser (Precision)",
                ["PDF Render Kalitesi"] = "PDF Render Quality",
                ["Render Kalitesi"] = "Render Quality",
                ["Sayfa netliği ve bellek kullanımı arasındaki dengeyi belirler."] = "Controls the balance between page sharpness and memory usage.",
                ["Hızlı"] = "Fast",
                ["Düşük RAM"] = "Low RAM",
                ["Dengeli"] = "Balanced",
                ["Önerilen"] = "Recommended",
                ["Ultra (4K)"] = "Ultra (4K)",
                ["Maksimum Netlik"] = "Max Clarity",
                ["Sayfa Okuma Filtresi"] = "Page Reading Filter",
                ["Uzun süreli çalışmalarda göz yorgunluğunu azaltmak için sayfa tonunu ayarlar."] = "Adjusts page tone to reduce eye strain during long study sessions.",
                ["Normal"] = "Normal",
                ["Orijinal Beyaz"] = "Original White",
                ["Sepya"] = "Sepia",
                ["Sıcak Ton"] = "Warm Tone",
                ["Gece / Koyu"] = "Dim / Night",
                ["Loş"] = "Dim",
                ["Önbellekte Tutulacak Sayfa Sayısı"] = "Cached Page Count",
                ["Önbellek Sayfa Sayısı"] = "Cached Page Count",
                ["Aktif sayfanın çevresinde yüksek çözünürlükte hazır tutulan sayfa sayısı."] = "Number of surrounding pages kept pre-rendered in high resolution.",
                ["Fare Tekerleği Kaydırma Hızı"] = "Mouse Wheel Scroll Speed",
                ["Kaydırma Hızı"] = "Scroll Speed",
                ["Sayfalar arasında dikey kaydırma yaparken kullanılan hız çarpanı."] = "Speed multiplier used when scrolling vertically through pages.",
                ["Varsayılan Yakınlaştırma (Zoom)"] = "Default Zoom Level",
                ["Varsayılan Yakınlaştırma"] = "Default Zoom",
                ["Yeni bir PDF açıldığında uygulanacak başlangıç yakınlaştırma oranı."] = "Initial zoom percentage applied when opening a new PDF.",
                ["Varsayılan Ek Çözüm Kağıdı"] = "Default Extra Solution Sheet",
                ["Ek Çözüm Kağıdı"] = "Extra Solution Sheet",
                ["PDF sayfalarının sağ tarafında ekstra işlem ve çözüm alanı açar."] = "Opens an extra scratchpad and solution area on the right side of pages.",
                ["Ek Çözüm Kağıdı Genişliği"] = "Extra Solution Sheet Width",
                ["Sağ taraftaki ek çözüm kağıdının piksel cinsinden genişliği."] = "Width of the right-side extra solution sheet in pixels.",
                ["Çözüm Kağıdı Deseni"] = "Solution Sheet Pattern",
                ["Kağıt Deseni"] = "Paper Pattern",
                ["Ek çözüm kağıdında kullanılacak varsayılan kılavuz çizgi deseni."] = "Default guide line pattern used on the extra solution sheet.",
                ["Boş Beyaz"] = "Blank White",
                ["Kareli"] = "Grid",
                ["Çizgili"] = "Ruled",
                ["Noktalı"] = "Dotted",
                ["Kılavuz Çizgi / Izgara Rengi"] = "Guide Line / Grid Color",
                ["Çizgi / Izgara Rengi"] = "Line / Grid Color",
                ["Kareli, çizgili veya noktalı kağıt deseninin çizgi rengi."] = "Line color for grid, ruled, or dotted paper patterns.",
                ["Koyu Gri"] = "Dark Gray",
                ["Açık Gri"] = "Light Gray",
                ["Klasik Mavi"] = "Classic Blue",
                ["Mavi"] = "Blue",
                ["İndigo"] = "Indigo",
                ["Toplam Kitap / PDF"] = "Total Books / PDFs",
                ["Çözümlü Sayfa"] = "Annotated Pages",
                ["Disk Kullanımı"] = "Disk Usage",
                ["Veri ve Kayıt Klasörü"] = "Data & Storage Folder",
                ["Kütüphane Klasörü"] = "Library Folder",
                ["Tüm çizimleriniz ve PDF verileriniz bu klasörde saklanır."] = "All your drawings and PDF metadata are stored in this folder.",
                ["Klasörü Aç"] = "Open Folder",
                ["Önbelleği Temizle"] = "Clear Cache",
                ["Geçici render belleğini temizleyerek RAM kullanımını sıfırlar."] = "Frees up RAM by clearing temporary rendered page bitmaps.",
                ["Tüm Çizimleri Sıfırla"] = "Reset All Drawings",
                ["Kayıtlı tüm sayfa çizimlerini kalıcı olarak siler."] = "Permanently deletes all saved page annotations and drawings.",
                ["Arayüz Dili"] = "Interface Language",
                ["Uygulama arayüzü ve menüler için tercih edilen dili seçin."] = "Select the preferred language for the application interface and menus.",
                ["Başlangıçta Açılış Davranışı"] = "Startup Behavior",
                ["SarpNote başlatıldığında ilk olarak hangi ekranın açılacağını seçin."] = "Choose which screen opens first when SarpNote starts.",
                ["Kitaplık Ekranını Aç"] = "Open Library Screen",
                ["Son Açılan PDF'den Devam Et"] = "Resume Last Opened PDF",
                ["Boş Çalışma Alanı Aç"] = "Open Blank Workspace",
                ["Varsayılan Başlangıç Aracı"] = "Default Startup Tool",
                ["Başlangıç Aracı"] = "Startup Tool",
                ["PDF açıldığında otomatik seçili gelecek çizim aracı."] = "Drawing tool automatically selected when a PDF is opened.",
                ["Kalem"] = "Pen",
                ["Fosforlu Kalem"] = "Highlighter",
                ["Silgi"] = "Eraser",
                ["El (Kaydırma)"] = "Hand (Pan)",
                ["Başlangıç Kalem Kalınlığı"] = "Startup Pen Thickness",
                ["Uygulama açıldığında kalemin başlayacağı varsayılan uç kalınlığı."] = "Default pen tip thickness when the application starts.",
                ["Başlangıç Silgi Boyutu"] = "Startup Eraser Size",
                ["Uygulama açıldığında silginin başlayacağı varsayılan boyut."] = "Default eraser size when the application starts.",
                ["Başlangıç Pencere Modu"] = "Startup Window Mode",
                ["Pencere Durumu"] = "Window State",
                ["Uygulamanın tam ekran mı yoksa normal pencerede mi açılacağını belirler."] = "Determines whether the app starts maximized or in a normal window.",
                ["Tam Ekran (Maksimum)"] = "Maximized (Full Screen)",
                ["Normal Pencere"] = "Normal Window",
                ["Pen Studio"] = "Pen Studio",
                ["Kalem Stüdyosu"] = "Pen Studio",
                ["Kalınlık"] = "Thickness",
                ["Renkler"] = "Colors",
                ["Opaklık"] = "Opacity",
                ["- - -  Kesikli Çizgi"] = "- - -  Dashed Line",
                ["➔  Vektör Oku"] = "➔  Vector Arrow",
                ["Hızlı Favori Kalemler (Sol Tık: Seç | Sağ Tık: Kaydet)"] = "Quick Favorite Pens (Left: Select | Right: Save)",
                ["◫ Cevap Anahtarı / Bölünmüş Sayfa"] = "◫ Answer Key / Split Page",
                ["Son"] = "Last",
                ["SarpNote - Kırpılan Soru"] = "SarpNote - Cropped Question",
                ["✓ Soru panoya kopyalandı (Ctrl+V ile yapıştırabilir veya aşağıdan kaydedebilirsin):"] = "✓ Question copied to clipboard (Paste with Ctrl+V or save below):",
                ["📌 Hata Defterine Kaydet"] = "📌 Save to Mistake Notebook",
                ["💾 Farklı Kaydet (PNG)..."] = "💾 Save As (PNG)..."
            },

            ["de"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PAGES"] = "SEITEN",
                ["SAYFALAR"] = "SEITEN",
                ["Sayfalar"] = "Seiten",
                ["Sayfa"] = "Seite",
                ["Tüm Kitaplar"] = "Alle Bücher",
                ["Yıldızlı Sorularım"] = "Markierte Fragen",
                ["Matematik & Geometri"] = "Mathematik & Geometrie",
                ["TYT / AYT Deneme"] = "TYT / AYT Übungstest",
                ["Kitap adı, klasör veya ders ara..."] = "Buchname, Ordner oder Fach suchen...",
                ["+ Kitap Ekle"] = "+ Buch hinzufügen",
                ["Kitap Ekle"] = "Buch hinzufügen",
                ["Çizime Dön"] = "Zurück zum Zeichnen",
                ["Kitap Adı"] = "Buchname",
                ["Klasör Konumu"] = "Ordnerpfad",
                ["Kaldığın Sayfa"] = "Letzte Seite",
                ["Son Açılma"] = "Zuletzt geöffnet",
                ["Ders"] = "Fach",
                ["Dersler"] = "Fächer",
                ["Kategoriler"] = "Kategorien",
                ["Bugün"] = "Heute",
                ["Dün"] = "Gestern",
                ["Az önce"] = "Gerade eben",
                ["Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran"] = "Auto-Lösungsblatt & Vollbild bei neuen PDFs",
                ["Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar."] = "Bereitet für jeden geöffneten Test automatisch das rechte Lösungsblatt vor.",
                ["Kapalı"] = "Aus",
                ["Açık"] = "Ein",
                ["Bölünmüş"] = "Geteilt",
                ["◫ Bölünmüş"] = "◫ Geteilt",
                ["Kırp"] = "Ausschneiden",
                ["✂ Kırp"] = "✂ Ausschneiden",
                ["Dışa Aktar"] = "Exportieren",
                ["📤 Dışa Aktar"] = "📤 Exportieren",
                ["Kitaplık"] = "Bibliothek",
                ["KÜTÜPHANE"] = "BIBLIOTHEK",
                ["KİTAPLIK"] = "BIBLIOTHEK",
                ["KLASÖRLER"] = "ORDNER",
                ["Klasörler"] = "Ordner",
                ["ETİKETLER"] = "TAGS",
                ["DERS ETİKETLERİ"] = "FÄCHER-TAGS",
                ["Ders Etiketleri"] = "Fächer-Tags",
                ["Etiketler"] = "Tags",
                ["Tüm Belgeler"] = "Alle Dokumente",
                ["Tüm PDF'ler"] = "Alle PDFs",
                ["Son Açılanlar"] = "Zuletzt geöffnet",
                ["Favoriler"] = "Favoriten",
                ["Yıldızlılar"] = "Markiert",
                ["PDF Ekle"] = "PDF hinzufügen",
                ["+ PDF Ekle"] = "+ PDF hinzufügen",
                ["Yeni Klasör"] = "Neuer Ordner",
                ["+ Yeni Klasör"] = "+ Neuer Ordner",
                ["+ Klasör"] = "+ Ordner",
                ["Klasör Oluştur"] = "Ordner erstellen",
                ["Yeni Not Defteri"] = "Neues Notizbuch",
                ["Kitaplıkta ara..."] = "In Bibliothek suchen...",
                ["Ara..."] = "Suchen...",
                ["Henüz PDF eklenmedi"] = "Noch keine PDFs hinzugefügt",
                ["Aç"] = "Öffnen",
                ["Sil"] = "Löschen",
                ["Kaldır"] = "Entfernen",
                ["Genel"] = "Allgemein",
                ["Matematik"] = "Mathematik",
                ["Geometri"] = "Geometrie",
                ["Fizik"] = "Physik",
                ["Kimya"] = "Chemie",
                ["Biyoloji"] = "Biologie",
                ["Türkçe"] = "Türkisch",
                ["Tarih"] = "Geschichte",
                ["Coğrafya"] = "Geografie",
                ["Deneme"] = "Übungstest",
                ["Tümü"] = "Alle",
                ["AYARLAR"] = "EINSTELLUNGEN",
                ["Ayarlar"] = "Einstellungen",
                ["Uygulama Ayarları"] = "Anwendungseinstellungen",
                ["Kaydet ve Kapat"] = "Speichern & Schließen",
                ["Tamam"] = "OK",
                ["İptal"] = "Abbrechen",
                ["Kapat"] = "Schließen",
                ["Varsayılana Sıfırla"] = "Auf Standard zurücksetzen",
                ["Varsayılanlara Sıfırla"] = "Auf Standard zurücksetzen",
                ["Sıfırla"] = "Zurücksetzen",
                ["Kalem & Silgi"] = "Stift & Radierer",
                ["Kalem ve Silgi"] = "Stift & Radierer",
                ["Kalem & Silgi Motoru"] = "Stift- & Radier-Engine",
                ["Kalem ve Silgi Motoru"] = "Stift- & Radier-Engine",
                ["Çizim hassasiyeti, şekil tanıma ve silgi davranışlarını özelleştirin."] = "Zeichenpräzision, Formerkennung und Radierverhalten anpassen.",
                ["Görüntü & Performans"] = "Anzeige & Leistung",
                ["Görüntü ve Performans"] = "Anzeige & Leistung",
                ["4K PDF render kalitesi, okuma filtreleri ve kaydırma akıcılığını ayarlayın."] = "4K-PDF-Renderqualität, Lesefilter und Scroll-Flüssigkeit anpassen.",
                ["Çalışma Alanı"] = "Arbeitsbereich",
                ["Çalışma Alanı & Kağıt"] = "Arbeitsbereich & Papier",
                ["Kağıt & Çalışma Alanı"] = "Papier & Arbeitsbereich",
                ["Varsayılan yakınlaştırma, ek çözüm kağıdı ve kılavuz çizgi tercihlerini yönetin."] = "Standard-Zoom, zusätzliches Lösungsblatt und Hilfslinien verwalten.",
                ["Depolama"] = "Speicher",
                ["Depolama & Veri"] = "Speicher & Daten",
                ["Depolama ve Veri Yönetimi"] = "Speicher- & Datenverwaltung",
                ["Kütüphane istatistiklerini görüntüleyin, önbelleği ve kayıt klasörünü yönetin."] = "Bibliotheksstatistiken anzeigen, Cache und Speicherordner verwalten.",
                ["Başlangıç"] = "Start",
                ["Başlangıç Ayarları"] = "Starteinstellungen",
                ["Uygulama açılışındaki varsayılan ekran, araç ve pencere tercihlerini belirleyin."] = "Standardbildschirm, Werkzeug, Sprache und Fenster beim Start festlegen.",
                ["Varsayılan Kalem Modu"] = "Standard-Stiftmodus",
                ["Çizim yaparken kullanılacak varsayılan kalem motorunu seçin."] = "Wählen Sie die Standard-Stift-Engine beim Zeichnen.",
                ["Basic"] = "Einfach",
                ["Pro"] = "Pro",
                ["Basit"] = "Einfach",
                ["Basic (Sabit Uç)"] = "Einfach (Feste Spitze)",
                ["Basit (Sabit Kalınlık)"] = "Einfach (Feste Stärke)",
                ["Basınçsız, sabit kalınlık"] = "Ohne Druck, feste Stärke",
                ["Pro (Basınç Duyarlı)"] = "Pro (Druckempfindlich)",
                ["Basınç destekli dinamik uç"] = "Druckempfindliche dynamische Spitze",
                ["Basılı Tutunca Şekle Dönüştür"] = "Halten für Formerkennung",
                ["Basılı tutunca otomatik şekil"] = "Halten für Auto-Form",
                ["Çizgi veya kapalı şekil çizip ucunda bekletince düzgün geometriye dönüştürür."] = "Wandelt Linien oder Formen beim Halten am Ende in exakte Geometrie um.",
                ["Basılı Tutma Süresi (Otomatik Şekil)"] = "Haltedauer (Auto-Form)",
                ["Basılı Tutma Süresi"] = "Haltedauer",
                ["Çizgiyi düzleştirmek veya şekle dönüştürmek için gereken bekleme süresi."] = "Wartezeit zum Begradigen von Linien oder Umwandeln von Formen.",
                ["Çizgi Yumuşatma (Smoothing)"] = "Strichglättung (Smoothing)",
                ["Çizgi Yumuşatma"] = "Strichglättung",
                ["El titremesini azaltır ve çizgilerin daha pürüzsüz görünmesini sağlar."] = "Reduziert Handzittern und sorgt für glattere Linien.",
                ["Varsayılan Silgi Modu"] = "Standard-Radiermodus",
                ["Silgi Modu"] = "Radiermodus",
                ["Silginin çizgileri tek dokunuşla mı yoksa bölgesel mi sileceğini belirler."] = "Legt fest, ob ganze Striche oder nur berührte Punkte gelöscht werden.",
                ["Çizgi Silgisi (Tüm Çizgiyi Siler)"] = "Strichradierer (Ganzen Strich löschen)",
                ["Nokta Silgisi (Kısmi Siler)"] = "Punktradierer (Teilweise löschen)",
                ["Nokta Silgisi (Hassas)"] = "Punktradierer (Präzise)",
                ["PDF Render Kalitesi"] = "PDF-Renderqualität",
                ["Render Kalitesi"] = "Renderqualität",
                ["Sayfa netliği ve bellek kullanımı arasındaki dengeyi belirler."] = "Steuert die Balance zwischen Seitenschärfe und Speicherverbrauch.",
                ["Hızlı"] = "Schnell",
                ["Düşük RAM"] = "Wenig RAM",
                ["Dengeli"] = "Ausgewogen",
                ["Önerilen"] = "Empfohlen",
                ["Ultra (4K)"] = "Ultra (4K)",
                ["Maksimum Netlik"] = "Maximale Schärfe",
                ["Sayfa Okuma Filtresi"] = "Seiten-Lesefilter",
                ["Uzun süreli çalışmalarda göz yorgunluğunu azaltmak için sayfa tonunu ayarlar."] = "Passt den Farbton der Seite an, um die Augen beim Lernen zu schonen.",
                ["Normal"] = "Normal",
                ["Orijinal Beyaz"] = "Original Weiß",
                ["Sepya"] = "Sepia",
                ["Sıcak Ton"] = "Warmer Farbton",
                ["Gece / Koyu"] = "Gedimmt / Nacht",
                ["Loş"] = "Gedimmt",
                ["Önbellekte Tutulacak Sayfa Sayısı"] = "Zwischengespeicherte Seiten",
                ["Önbellek Sayfa Sayısı"] = "Zwischengespeicherte Seiten",
                ["Aktif sayfanın çevresinde yüksek çözünürlükte hazır tutulan sayfa sayısı."] = "Anzahl umliegender Seiten, die hochauflösend im Speicher gehalten werden.",
                ["Fare Tekerleği Kaydırma Hızı"] = "Mausrad-Scrollgeschwindigkeit",
                ["Kaydırma Hızı"] = "Scrollgeschwindigkeit",
                ["Sayfalar arasında dikey kaydırma yaparken kullanılan hız çarpanı."] = "Geschwindigkeitsfaktor beim vertikalen Scrollen durch Seiten.",
                ["Varsayılan Yakınlaştırma (Zoom)"] = "Standard-Zoomstufe",
                ["Varsayılan Yakınlaştırma"] = "Standard-Zoom",
                ["Yeni bir PDF açıldığında uygulanacak başlangıç yakınlaştırma oranı."] = "Anfängliche Zoomstufe beim Öffnen eines neuen PDFs.",
                ["Varsayılan Ek Çözüm Kağıdı"] = "Standard-Lösungsblatt",
                ["Ek Çözüm Kağıdı"] = "Zusätzliches Lösungsblatt",
                ["PDF sayfalarının sağ tarafında ekstra işlem ve çözüm alanı açar."] = "Öffnet rechts neben den PDF-Seiten einen zusätzlichen Lösungsbereich.",
                ["Ek Çözüm Kağıdı Genişliği"] = "Breite des Lösungsblatts",
                ["Sağ taraftaki ek çözüm kağıdının piksel cinsinden genişliği."] = "Breite des rechten zusätzlichen Lösungsblatts in Pixeln.",
                ["Çözüm Kağıdı Deseni"] = "Lösungsblatt-Muster",
                ["Kağıt Deseni"] = "Papiermuster",
                ["Ek çözüm kağıdında kullanılacak varsayılan kılavuz çizgi deseni."] = "Standard-Hilfslinienmuster für das zusätzliche Lösungsblatt.",
                ["Boş Beyaz"] = "Blanko Weiß",
                ["Kareli"] = "Kariert",
                ["Çizgili"] = "Liniert",
                ["Noktalı"] = "Punktiert",
                ["Kılavuz Çizgi / Izgara Rengi"] = "Hilfslinien- / Rasterfarbe",
                ["Çizgi / Izgara Rengi"] = "Linien- / Rasterfarbe",
                ["Kareli, çizgili veya noktalı kağıt deseninin çizgi rengi."] = "Linienfarbe für kariertes, liniertes oder punktiertes Papier.",
                ["Koyu Gri"] = "Dunkelgrau",
                ["Açık Gri"] = "Hellgrau",
                ["Klasik Mavi"] = "Klassisch Blau",
                ["Mavi"] = "Blau",
                ["İndigo"] = "Indigo",
                ["Toplam Kitap / PDF"] = "Bücher / PDFs gesamt",
                ["Çözümlü Sayfa"] = "Bearbeitete Seiten",
                ["Disk Kullanımı"] = "Speichernutzung",
                ["Veri ve Kayıt Klasörü"] = "Daten- & Speicherordner",
                ["Kütüphane Klasörü"] = "Bibliotheksordner",
                ["Tüm çizimleriniz ve PDF verileriniz bu klasörde saklanır."] = "Alle Zeichnungen und PDF-Metadaten werden hier gespeichert.",
                ["Klasörü Aç"] = "Ordner öffnen",
                ["Önbelleği Temizle"] = "Cache leeren",
                ["Geçici render belleğini temizleyerek RAM kullanımını sıfırlar."] = "Gibt RAM frei, indem temporäre Seitenbilder geleert werden.",
                ["Tüm Çizimleri Sıfırla"] = "Alle Zeichnungen zurücksetzen",
                ["Kayıtlı tüm sayfa çizimlerini kalıcı olarak siler."] = "Löscht dauerhaft alle gespeicherten Seitenzeichnungen.",
                ["Arayüz Dili"] = "Oberflächensprache",
                ["Uygulama arayüzü ve menüler için tercih edilen dili seçin."] = "Wählen Sie die bevorzugte Sprache für Benutzeroberfläche und Menüs.",
                ["Başlangıçta Açılış Davranışı"] = "Startverhalten",
                ["SarpNote başlatıldığında ilk olarak hangi ekranın açılacağını seçin."] = "Wählen Sie, welcher Bildschirm beim Start von SarpNote geöffnet wird.",
                ["Kitaplık Ekranını Aç"] = "Bibliothek öffnen",
                ["Son Açılan PDF'den Devam Et"] = "Letztes PDF fortsetzen",
                ["Boş Çalışma Alanı Aç"] = "Leeren Arbeitsbereich öffnen",
                ["Varsayılan Başlangıç Aracı"] = "Standard-Startwerkzeug",
                ["Başlangıç Aracı"] = "Startwerkzeug",
                ["PDF açıldığında otomatik seçili gelecek çizim aracı."] = "Zeichenwerkzeug, das beim Öffnen eines PDFs automatisch aktiv ist.",
                ["Kalem"] = "Stift",
                ["Fosforlu Kalem"] = "Textmarker",
                ["Silgi"] = "Radierer",
                ["El (Kaydırma)"] = "Hand (Verschieben)",
                ["Başlangıç Kalem Kalınlığı"] = "Start-Stiftstärke",
                ["Uygulama açıldığında kalemin başlayacağı varsayılan uç kalınlığı."] = "Standard-Stiftspitzenstärke beim Start der Anwendung.",
                ["Başlangıç Silgi Boyutu"] = "Start-Radierergröße",
                ["Uygulama açıldığında silginin başlayacağı varsayılan boyut."] = "Standardgröße des Radierers beim Start der Anwendung.",
                ["Başlangıç Pencere Modu"] = "Start-Fenstermodus",
                ["Pencere Durumu"] = "Fensterstatus",
                ["Uygulamanın tam ekran mı yoksa normal pencerede mi açılacağını belirler."] = "Legt fest, ob die App maximiert oder im normalen Fenster startet.",
                ["Tam Ekran (Maksimum)"] = "Maximiert (Vollbild)",
                ["Normal Pencere"] = "Normales Fenster",
                ["Pen Studio"] = "Stift-Studio",
                ["Kalem Stüdyosu"] = "Stift-Studio",
                ["Kalınlık"] = "Stärke",
                ["Renkler"] = "Farben",
                ["Opaklık"] = "Deckkraft",
                ["- - -  Kesikli Çizgi"] = "- - -  Gestrichelte Linie",
                ["➔  Vektör Oku"] = "➔  Vektorpfeil",
                ["Hızlı Favori Kalemler (Sol Tık: Seç | Sağ Tık: Kaydet)"] = "Favoritenstifte (Links: Wählen | Rechts: Speichern)",
                ["◫ Cevap Anahtarı / Bölünmüş Sayfa"] = "◫ Lösungsschlüssel / Geteilte Seite",
                ["Son"] = "Ende",
                ["SarpNote - Kırpılan Soru"] = "SarpNote - Ausgeschnittene Aufgabe",
                ["✓ Soru panoya kopyalandı (Ctrl+V ile yapıştırabilir veya aşağıdan kaydedebilirsin):"] = "✓ Aufgabe in Zwischenablage kopiert (Strg+V zum Einfügen oder unten speichern):",
                ["📌 Hata Defterine Kaydet"] = "📌 Im Fehlerheft speichern",
                ["💾 Farklı Kaydet (PNG)..."] = "💾 Speichern unter (PNG)..."
            },

            ["fr"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PAGES"] = "PAGES",
                ["SAYFALAR"] = "PAGES",
                ["Sayfalar"] = "Pages",
                ["Sayfa"] = "Page",
                ["Tüm Kitaplar"] = "Tous les livres",
                ["Yıldızlı Sorularım"] = "Questions favorites",
                ["Matematik & Geometri"] = "Mathématiques & Géométrie",
                ["TYT / AYT Deneme"] = "Examen blanc TYT / AYT",
                ["Kitap adı, klasör veya ders ara..."] = "Rechercher un livre, dossier ou matière...",
                ["+ Kitap Ekle"] = "+ Ajouter un livre",
                ["Kitap Ekle"] = "Ajouter un livre",
                ["Çizime Dön"] = "Retour au dessin",
                ["Kitap Adı"] = "Nom du livre",
                ["Klasör Konumu"] = "Emplacement du dossier",
                ["Kaldığın Sayfa"] = "Dernière page",
                ["Son Açılma"] = "Ouvert dernièrement",
                ["Ders"] = "Matière",
                ["Dersler"] = "Matières",
                ["Kategoriler"] = "Catégories",
                ["Bugün"] = "Aujourd'hui",
                ["Dün"] = "Hier",
                ["Az önce"] = "À l'instant",
                ["Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran"] = "Feuille de brouillon & Plein écran auto",
                ["Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar."] = "Prépare automatiquement la feuille de calcul à droite pour chaque test.",
                ["Kapalı"] = "Désactivé",
                ["Açık"] = "Activé",
                ["Bölünmüş"] = "Écran scindé",
                ["◫ Bölünmüş"] = "◫ Écran scindé",
                ["Kırp"] = "Rogner",
                ["✂ Kırp"] = "✂ Rogner",
                ["Dışa Aktar"] = "Exporter",
                ["📤 Dışa Aktar"] = "📤 Exporter",
                ["Kitaplık"] = "Bibliothèque",
                ["KÜTÜPHANE"] = "BIBLIOTHÈQUE",
                ["KİTAPLIK"] = "BIBLIOTHÈQUE",
                ["KLASÖRLER"] = "DOSSIERS",
                ["Klasörler"] = "Dossiers",
                ["ETİKETLER"] = "ÉTIQUETTES",
                ["DERS ETİKETLERİ"] = "MATIÈRES",
                ["Ders Etiketleri"] = "Matières",
                ["Etiketler"] = "Étiquettes",
                ["Tüm Belgeler"] = "Tous les documents",
                ["Tüm PDF'ler"] = "Tous les PDF",
                ["Son Açılanlar"] = "Récents",
                ["Favoriler"] = "Favoris",
                ["Yıldızlılar"] = "Favoris",
                ["PDF Ekle"] = "Ajouter un PDF",
                ["+ PDF Ekle"] = "+ Ajouter un PDF",
                ["Yeni Klasör"] = "Nouveau dossier",
                ["+ Yeni Klasör"] = "+ Nouveau dossier",
                ["+ Klasör"] = "+ Dossier",
                ["Klasör Oluştur"] = "Créer un dossier",
                ["Yeni Not Defteri"] = "Nouveau cahier",
                ["Kitaplıkta ara..."] = "Rechercher dans la bibliothèque...",
                ["Ara..."] = "Rechercher...",
                ["Henüz PDF eklenmedi"] = "Aucun PDF ajouté pour le moment",
                ["Aç"] = "Ouvrir",
                ["Sil"] = "Supprimer",
                ["Kaldır"] = "Retirer",
                ["Genel"] = "Général",
                ["Matematik"] = "Mathématiques",
                ["Geometri"] = "Géométrie",
                ["Fizik"] = "Physique",
                ["Kimya"] = "Chimie",
                ["Biyoloji"] = "Biologie",
                ["Türkçe"] = "Turc",
                ["Tarih"] = "Histoire",
                ["Coğrafya"] = "Géographie",
                ["Deneme"] = "Examen blanc",
                ["Tümü"] = "Tout",
                ["AYARLAR"] = "PARAMÈTRES",
                ["Ayarlar"] = "Paramètres",
                ["Uygulama Ayarları"] = "Paramètres de l'application",
                ["Kaydet ve Kapat"] = "Enregistrer et fermer",
                ["Tamam"] = "OK",
                ["İptal"] = "Annuler",
                ["Kapat"] = "Fermer",
                ["Varsayılana Sıfırla"] = "Réinitialiser par défaut",
                ["Varsayılanlara Sıfırla"] = "Réinitialiser par défaut",
                ["Sıfırla"] = "Réinitialiser",
                ["Kalem & Silgi"] = "Stylo & Gomme",
                ["Kalem ve Silgi"] = "Stylo & Gomme",
                ["Kalem & Silgi Motoru"] = "Moteur Stylo & Gomme",
                ["Kalem ve Silgi Motoru"] = "Moteur Stylo & Gomme",
                ["Çizim hassasiyeti, şekil tanıma ve silgi davranışlarını özelleştirin."] = "Personnalisez la précision du dessin, la détection de formes et la gomme.",
                ["Görüntü & Performans"] = "Affichage & Performances",
                ["Görüntü ve Performans"] = "Affichage & Performances",
                ["4K PDF render kalitesi, okuma filtreleri ve kaydırma akıcılığını ayarlayın."] = "Ajustez la qualité de rendu 4K, les filtres de lecture et la fluidité.",
                ["Çalışma Alanı"] = "Espace de travail",
                ["Çalışma Alanı & Kağıt"] = "Espace de travail & Papier",
                ["Kağıt & Çalışma Alanı"] = "Papier & Espace de travail",
                ["Varsayılan yakınlaştırma, ek çözüm kağıdı ve kılavuz çizgi tercihlerini yönetin."] = "Gérez le zoom par défaut, la feuille de brouillon et le quadrillage.",
                ["Depolama"] = "Stockage",
                ["Depolama & Veri"] = "Stockage & Données",
                ["Depolama ve Veri Yönetimi"] = "Gestion du stockage et des données",
                ["Kütüphane istatistiklerini görüntüleyin, önbelleği ve kayıt klasörünü yönetin."] = "Consultez les statistiques, gérez le cache et le dossier de stockage.",
                ["Başlangıç"] = "Démarrage",
                ["Başlangıç Ayarları"] = "Paramètres de démarrage",
                ["Uygulama açılışındaki varsayılan ekran, araç ve pencere tercihlerini belirleyin."] = "Définissez l'écran, l'outil, la langue et la fenêtre au démarrage.",
                ["Varsayılan Kalem Modu"] = "Mode de stylo par défaut",
                ["Çizim yaparken kullanılacak varsayılan kalem motorunu seçin."] = "Sélectionnez le moteur de stylo utilisé par défaut.",
                ["Basic"] = "Simple",
                ["Pro"] = "Pro",
                ["Basit"] = "Simple",
                ["Basic (Sabit Uç)"] = "Simple (Pointe fixe)",
                ["Basit (Sabit Kalınlık)"] = "Simple (Épaisseur fixe)",
                ["Basınçsız, sabit kalınlık"] = "Sans pression, épaisseur fixe",
                ["Pro (Basınç Duyarlı)"] = "Pro (Sensible à la pression)",
                ["Basınç destekli dinamik uç"] = "Pointe dynamique sensible à la pression",
                ["Basılı Tutunca Şekle Dönüştür"] = "Maintenir pour former",
                ["Basılı tutunca otomatik şekil"] = "Maintenir pour forme auto",
                ["Çizgi veya kapalı şekil çizip ucunda bekletince düzgün geometriye dönüştürür."] = "Convertit les lignes ou formes maintenues en géométrie parfaite.",
                ["Basılı Tutma Süresi (Otomatik Şekil)"] = "Durée de maintien (Forme auto)",
                ["Basılı Tutma Süresi"] = "Durée de maintien",
                ["Çizgiyi düzleştirmek veya şekle dönüştürmek için gereken bekleme süresi."] = "Temps d'attente requis pour redresser une ligne ou convertir une forme.",
                ["Çizgi Yumuşatma (Smoothing)"] = "Lissage du trait (Smoothing)",
                ["Çizgi Yumuşatma"] = "Lissage du trait",
                ["El titremesini azaltır ve çizgilerin daha pürüzsüz görünmesini sağlar."] = "Réduit les tremblements et rend les traits plus fluides.",
                ["Varsayılan Silgi Modu"] = "Mode de gomme par défaut",
                ["Silgi Modu"] = "Mode de gomme",
                ["Silginin çizgileri tek dokunuşla mı yoksa bölgesel mi sileceğini belirler."] = "Détermine si la gomme efface tout le trait ou seulement la zone touchée.",
                ["Çizgi Silgisi (Tüm Çizgiyi Siler)"] = "Gomme de trait (Efface tout le trait)",
                ["Nokta Silgisi (Kısmi Siler)"] = "Gomme de précision (Effacement partiel)",
                ["Nokta Silgisi (Hassas)"] = "Gomme de précision",
                ["PDF Render Kalitesi"] = "Qualité de rendu PDF",
                ["Render Kalitesi"] = "Qualité de rendu",
                ["Sayfa netliği ve bellek kullanımı arasındaki dengeyi belirler."] = "Contrôle l'équilibre entre la netteté des pages et la mémoire.",
                ["Hızlı"] = "Rapide",
                ["Düşük RAM"] = "RAM faible",
                ["Dengeli"] = "Équilibré",
                ["Önerilen"] = "Recommandé",
                ["Ultra (4K)"] = "Ultra (4K)",
                ["Maksimum Netlik"] = "Clarté maximale",
                ["Sayfa Okuma Filtresi"] = "Filtre de lecture des pages",
                ["Uzun süreli çalışmalarda göz yorgunluğunu azaltmak için sayfa tonunu ayarlar."] = "Ajuste la teinte de la page pour réduire la fatigue oculaire.",
                ["Normal"] = "Normal",
                ["Orijinal Beyaz"] = "Blanc original",
                ["Sepya"] = "Sépia",
                ["Sıcak Ton"] = "Ton chaud",
                ["Gece / Koyu"] = "Tamisé / Nuit",
                ["Loş"] = "Tamisé",
                ["Önbellekte Tutulacak Sayfa Sayısı"] = "Nombre de pages en cache",
                ["Önbellek Sayfa Sayısı"] = "Pages en cache",
                ["Aktif sayfanın çevresinde yüksek çözünürlükte hazır tutulan sayfa sayısı."] = "Nombre de pages voisines pré-rendues en haute résolution.",
                ["Fare Tekerleği Kaydırma Hızı"] = "Vitesse de défilement de la molette",
                ["Kaydırma Hızı"] = "Vitesse de défilement",
                ["Sayfalar arasında dikey kaydırma yaparken kullanılan hız çarpanı."] = "Multiplicateur de vitesse lors du défilement vertical des pages.",
                ["Varsayılan Yakınlaştırma (Zoom)"] = "Niveau de zoom par défaut",
                ["Varsayılan Yakınlaştırma"] = "Zoom par défaut",
                ["Yeni bir PDF açıldığında uygulanacak başlangıç yakınlaştırma oranı."] = "Pourcentage de zoom initial appliqué à l'ouverture d'un PDF.",
                ["Varsayılan Ek Çözüm Kağıdı"] = "Feuille de résolution par défaut",
                ["Ek Çözüm Kağıdı"] = "Feuille de résolution latérale",
                ["PDF sayfalarının sağ tarafında ekstra işlem ve çözüm alanı açar."] = "Ouvre un espace de brouillon et de calcul à droite des pages PDF.",
                ["Ek Çözüm Kağıdı Genişliği"] = "Largeur de la feuille latérale",
                ["Sağ taraftaki ek çözüm kağıdının piksel cinsinden genişliği."] = "Largeur en pixels de la feuille de résolution située à droite.",
                ["Çözüm Kağıdı Deseni"] = "Modèle de la feuille latérale",
                ["Kağıt Deseni"] = "Modèle de papier",
                ["Ek çözüm kağıdında kullanılacak varsayılan kılavuz çizgi deseni."] = "Modèle de quadrillage par défaut pour la feuille latérale.",
                ["Boş Beyaz"] = "Blanc vierge",
                ["Kareli"] = "Quadrillé",
                ["Çizgili"] = "Ligné",
                ["Noktalı"] = "Pointillé",
                ["Kılavuz Çizgi / Izgara Rengi"] = "Couleur du quadrillage",
                ["Çizgi / Izgara Rengi"] = "Couleur de grille / ligne",
                ["Kareli, çizgili veya noktalı kağıt deseninin çizgi rengi."] = "Couleur des lignes pour le papier quadrillé, ligné ou pointillé.",
                ["Koyu Gri"] = "Gris foncé",
                ["Açık Gri"] = "Gris clair",
                ["Klasik Mavi"] = "Bleu classique",
                ["Mavi"] = "Bleu",
                ["İndigo"] = "Indigo",
                ["Toplam Kitap / PDF"] = "Total Livres / PDF",
                ["Çözümlü Sayfa"] = "Pages annotées",
                ["Disk Kullanımı"] = "Espace disque utilisé",
                ["Veri ve Kayıt Klasörü"] = "Dossier de données et stockage",
                ["Kütüphane Klasörü"] = "Dossier de la bibliothèque",
                ["Tüm çizimleriniz ve PDF verileriniz bu klasörde saklanır."] = "Tous vos tracés et métadonnées PDF sont stockés dans ce dossier.",
                ["Klasörü Aç"] = "Ouvrir le dossier",
                ["Önbelleği Temizle"] = "Vider le cache",
                ["Geçici render belleğini temizleyerek RAM kullanımını sıfırlar."] = "Libère la RAM en vidant le cache des pages temporaires.",
                ["Tüm Çizimleri Sıfırla"] = "Réinitialiser tous les tracés",
                ["Kayıtlı tüm sayfa çizimlerini kalıcı olarak siler."] = "Supprime définitivement toutes les annotations enregistrées.",
                ["Arayüz Dili"] = "Langue de l'interface",
                ["Uygulama arayüzü ve menüler için tercih edilen dili seçin."] = "Sélectionnez la langue préférée pour l'interface et les menus.",
                ["Başlangıçta Açılış Davranışı"] = "Comportement au démarrage",
                ["SarpNote başlatıldığında ilk olarak hangi ekranın açılacağını seçin."] = "Choisissez quel écran s'ouvre en premier au lancement de SarpNote.",
                ["Kitaplık Ekranını Aç"] = "Ouvrir la bibliothèque",
                ["Son Açılan PDF'den Devam Et"] = "Reprendre le dernier PDF ouvert",
                ["Boş Çalışma Alanı Aç"] = "Ouvrir un espace de travail vierge",
                ["Varsayılan Başlangıç Aracı"] = "Outil de démarrage par défaut",
                ["Başlangıç Aracı"] = "Outil de démarrage",
                ["PDF açıldığında otomatik seçili gelecek çizim aracı."] = "Outil de dessin sélectionné automatiquement à l'ouverture d'un PDF.",
                ["Kalem"] = "Stylo",
                ["Fosforlu Kalem"] = "Surligneur",
                ["Silgi"] = "Gomme",
                ["El (Kaydırma)"] = "Main (Déplacer)",
                ["Başlangıç Kalem Kalınlığı"] = "Épaisseur initiale du stylo",
                ["Uygulama açıldığında kalemin başlayacağı varsayılan uç kalınlığı."] = "Épaisseur par défaut de la pointe du stylo au démarrage.",
                ["Başlangıç Silgi Boyutu"] = "Taille initiale de la gomme",
                ["Uygulama açıldığında silginin başlayacağı varsayılan boyut."] = "Taille par défaut de la gomme au démarrage de l'application.",
                ["Başlangıç Pencere Modu"] = "Mode de fenêtre au démarrage",
                ["Pencere Durumu"] = "État de la fenêtre",
                ["Uygulamanın tam ekran mı yoksa normal pencerede mi açılacağını belirler."] = "Détermine si l'application démarre en plein écran ou en fenêtre normale.",
                ["Tam Ekran (Maksimum)"] = "Plein écran (Maximisé)",
                ["Normal Pencere"] = "Fenêtre normale",
                ["Pen Studio"] = "Studio Stylo",
                ["Kalem Stüdyosu"] = "Studio Stylo",
                ["Kalınlık"] = "Épaisseur",
                ["Renkler"] = "Couleurs",
                ["Opaklık"] = "Opacité",
                ["- - -  Kesikli Çizgi"] = "- - -  Ligne pointillée",
                ["➔  Vektör Oku"] = "➔  Flèche vectorielle",
                ["Hızlı Favori Kalemler (Sol Tık: Seç | Sağ Tık: Kaydet)"] = "Stylos favoris (Clic gauche : Choisir | Clic droit : Sauver)",
                ["◫ Cevap Anahtarı / Bölünmüş Sayfa"] = "◫ Corrigé / Page scindée",
                ["Son"] = "Fin",
                ["SarpNote - Kırpılan Soru"] = "SarpNote - Question rognée",
                ["✓ Soru panoya kopyalandı (Ctrl+V ile yapıştırabilir veya aşağıdan kaydedebilirsin):"] = "✓ Question copiée dans le presse-papiers (Collez avec Ctrl+V ou enregistrez ci-dessous) :",
                ["📌 Hata Defterine Kaydet"] = "📌 Enregistrer dans le carnet d'erreurs",
                ["💾 Farklı Kaydet (PNG)..."] = "💾 Enregistrer sous (PNG)..."
            },

            ["es"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["PAGES"] = "PÁGINAS",
                ["SAYFALAR"] = "PÁGINAS",
                ["Sayfalar"] = "Páginas",
                ["Sayfa"] = "Página",
                ["Tüm Kitaplar"] = "Todos los libros",
                ["Yıldızlı Sorularım"] = "Preguntas destacadas",
                ["Matematik & Geometri"] = "Matemáticas y Geometría",
                ["TYT / AYT Deneme"] = "Examen de práctica TYT / AYT",
                ["Kitap adı, klasör veya ders ara..."] = "Buscar libro, carpeta o asignatura...",
                ["+ Kitap Ekle"] = "+ Añadir libro",
                ["Kitap Ekle"] = "Añadir libro",
                ["Çizime Dön"] = "Volver al dibujo",
                ["Kitap Adı"] = "Nombre del libro",
                ["Klasör Konumu"] = "Ubicación de carpeta",
                ["Kaldığın Sayfa"] = "Última página",
                ["Son Açılma"] = "Última apertura",
                ["Ders"] = "Asignatura",
                ["Dersler"] = "Asignaturas",
                ["Kategoriler"] = "Categorías",
                ["Bugün"] = "Hoy",
                ["Dün"] = "Ayer",
                ["Az önce"] = "Ahora mismo",
                ["Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran"] = "Hoja de resolución y pantalla completa auto",
                ["Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar."] = "Prepara automáticamente la hoja lateral de operaciones en cada prueba.",
                ["Kapalı"] = "Desactivado",
                ["Açık"] = "Activado",
                ["Bölünmüş"] = "Vista dividida",
                ["◫ Bölünmüş"] = "◫ Vista dividida",
                ["Kırp"] = "Recortar",
                ["✂ Kırp"] = "✂ Recortar",
                ["Dışa Aktar"] = "Exportar",
                ["📤 Dışa Aktar"] = "📤 Exportar",
                ["Kitaplık"] = "Biblioteca",
                ["KÜTÜPHANE"] = "BIBLIOTECA",
                ["KİTAPLIK"] = "BIBLIOTECA",
                ["KLASÖRLER"] = "CARPETAS",
                ["Klasörler"] = "Carpetas",
                ["ETİKETLER"] = "ETIQUETAS",
                ["DERS ETİKETLERİ"] = "ASIGNATURAS",
                ["Ders Etiketleri"] = "Asignaturas",
                ["Etiketler"] = "Etiquetas",
                ["Tüm Belgeler"] = "Todos los documentos",
                ["Tüm PDF'ler"] = "Todos los PDF",
                ["Son Açılanlar"] = "Recientes",
                ["Favoriler"] = "Favoritos",
                ["Yıldızlılar"] = "Destacados",
                ["PDF Ekle"] = "Añadir PDF",
                ["+ PDF Ekle"] = "+ Añadir PDF",
                ["Yeni Klasör"] = "Nueva carpeta",
                ["+ Yeni Klasör"] = "+ Nueva carpeta",
                ["+ Klasör"] = "+ Carpeta",
                ["Klasör Oluştur"] = "Crear carpeta",
                ["Yeni Not Defteri"] = "Nuevo cuaderno",
                ["Kitaplıkta ara..."] = "Buscar en la biblioteca...",
                ["Ara..."] = "Buscar...",
                ["Henüz PDF eklenmedi"] = "Aún no se han añadido PDF",
                ["Aç"] = "Abrir",
                ["Sil"] = "Eliminar",
                ["Kaldır"] = "Quitar",
                ["Genel"] = "General",
                ["Matematik"] = "Matemáticas",
                ["Geometri"] = "Geometría",
                ["Fizik"] = "Física",
                ["Kimya"] = "Química",
                ["Biyoloji"] = "Biología",
                ["Türkçe"] = "Turco",
                ["Tarih"] = "Historia",
                ["Coğrafya"] = "Geografía",
                ["Deneme"] = "Examen de práctica",
                ["Tümü"] = "Todo",
                ["AYARLAR"] = "AJUSTES",
                ["Ayarlar"] = "Ajustes",
                ["Uygulama Ayarları"] = "Ajustes de la aplicación",
                ["Kaydet ve Kapat"] = "Guardar y cerrar",
                ["Tamam"] = "Aceptar",
                ["İptal"] = "Cancelar",
                ["Kapat"] = "Cerrar",
                ["Varsayılana Sıfırla"] = "Restablecer valores",
                ["Varsayılanlara Sıfırla"] = "Restablecer valores",
                ["Sıfırla"] = "Restablecer",
                ["Kalem & Silgi"] = "Lápiz y Borrador",
                ["Kalem ve Silgi"] = "Lápiz y Borrador",
                ["Kalem & Silgi Motoru"] = "Motor de Lápiz y Borrador",
                ["Kalem ve Silgi Motoru"] = "Motor de Lápiz y Borrador",
                ["Çizim hassasiyeti, şekil tanıma ve silgi davranışlarını özelleştirin."] = "Personalice la precisión de dibujo, detección de formas y el borrador.",
                ["Görüntü & Performans"] = "Pantalla y Rendimiento",
                ["Görüntü ve Performans"] = "Pantalla y Rendimiento",
                ["4K PDF render kalitesi, okuma filtreleri ve kaydırma akıcılığını ayarlayın."] = "Ajuste la calidad 4K del PDF, los filtros de lectura y el desplazamiento.",
                ["Çalışma Alanı"] = "Espacio de trabajo",
                ["Çalışma Alanı & Kağıt"] = "Espacio de trabajo y Papel",
                ["Kağıt & Çalışma Alanı"] = "Papel y Espacio de trabajo",
                ["Varsayılan yakınlaştırma, ek çözüm kağıdı ve kılavuz çizgi tercihlerini yönetin."] = "Gestione el zoom predeterminado, la hoja lateral y las líneas guía.",
                ["Depolama"] = "Almacenamiento",
                ["Depolama & Veri"] = "Almacenamiento y Datos",
                ["Depolama ve Veri Yönetimi"] = "Gestión de Almacenamiento y Datos",
                ["Kütüphane istatistiklerini görüntüleyin, önbelleği ve kayıt klasörünü yönetin."] = "Vea estadísticas de la biblioteca, gestione la caché y la carpeta de datos.",
                ["Başlangıç"] = "Inicio",
                ["Başlangıç Ayarları"] = "Ajustes de inicio",
                ["Uygulama açılışındaki varsayılan ekran, araç ve pencere tercihlerini belirleyin."] = "Defina la pantalla, herramienta, idioma y ventana al iniciar la aplicación.",
                ["Varsayılan Kalem Modu"] = "Modo de lápiz predeterminado",
                ["Çizim yaparken kullanılacak varsayılan kalem motorunu seçin."] = "Seleccione el motor de lápiz predeterminado al dibujar.",
                ["Basic"] = "Básico",
                ["Pro"] = "Pro",
                ["Basit"] = "Básico",
                ["Basic (Sabit Uç)"] = "Básico (Punta fija)",
                ["Basit (Sabit Kalınlık)"] = "Básico (Grosor fijo)",
                ["Basınçsız, sabit kalınlık"] = "Sin presión, grosor fijo",
                ["Pro (Basınç Duyarlı)"] = "Pro (Sensible a la presión)",
                ["Basınç destekli dinamik uç"] = "Punta dinámica sensible a la presión",
                ["Basılı Tutunca Şekle Dönüştür"] = "Mantener para formar figura",
                ["Basılı tutunca otomatik şekil"] = "Mantener para auto-forma",
                ["Çizgi veya kapalı şekil çizip ucunda bekletince düzgün geometriye dönüştürür."] = "Convierte líneas o figuras cerradas en geometría perfecta al mantener al final.",
                ["Basılı Tutma Süresi (Otomatik Şekil)"] = "Tiempo de espera (Auto-forma)",
                ["Basılı Tutma Süresi"] = "Tiempo de espera",
                ["Çizgiyi düzleştirmek veya şekle dönüştürmek için gereken bekleme süresi."] = "Tiempo de espera necesario para enderezar una línea o convertir una figura.",
                ["Çizgi Yumuşatma (Smoothing)"] = "Suavizado de trazo (Smoothing)",
                ["Çizgi Yumuşatma"] = "Suavizado de trazo",
                ["El titremesini azaltır ve çizgilerin daha pürüzsüz görünmesini sağlar."] = "Reduce el temblor de la mano y hace que los trazos sean más suaves.",
                ["Varsayılan Silgi Modu"] = "Modo de borrador predeterminado",
                ["Silgi Modu"] = "Modo de borrador",
                ["Silginin çizgileri tek dokunuşla mı yoksa bölgesel mi sileceğini belirler."] = "Determina si el borrador elimina trazos completos o solo puntos tocados.",
                ["Çizgi Silgisi (Tüm Çizgiyi Siler)"] = "Borrador de trazo (Borra todo el trazo)",
                ["Nokta Silgisi (Kısmi Siler)"] = "Borrador de punto (Borrado parcial)",
                ["Nokta Silgisi (Hassas)"] = "Borrador de punto (Preciso)",
                ["PDF Render Kalitesi"] = "Calidad de renderizado PDF",
                ["Render Kalitesi"] = "Calidad de renderizado",
                ["Sayfa netliği ve bellek kullanımı arasındaki dengeyi belirler."] = "Controla el equilibrio entre la nitidez de página y el uso de memoria.",
                ["Hızlı"] = "Rápido",
                ["Düşük RAM"] = "Baja RAM",
                ["Dengeli"] = "Equilibrado",
                ["Önerilen"] = "Recomendado",
                ["Ultra (4K)"] = "Ultra (4K)",
                ["Maksimum Netlik"] = "Máxima nitidez",
                ["Sayfa Okuma Filtresi"] = "Filtro de lectura de página",
                ["Uzun süreli çalışmalarda göz yorgunluğunu azaltmak için sayfa tonunu ayarlar."] = "Ajusta el tono de la página para reducir la fatiga visual al estudiar.",
                ["Normal"] = "Normal",
                ["Orijinal Beyaz"] = "Blanco original",
                ["Sepya"] = "Sepia",
                ["Sıcak Ton"] = "Tono cálido",
                ["Gece / Koyu"] = "Atenuado / Noche",
                ["Loş"] = "Atenuado",
                ["Önbellekte Tutulacak Sayfa Sayısı"] = "Cantidad de páginas en caché",
                ["Önbellek Sayfa Sayısı"] = "Páginas en caché",
                ["Aktif sayfanın çevresinde yüksek çözünürlükte hazır tutulan sayfa sayısı."] = "Número de páginas cercanas pre-renderizadas en alta resolución.",
                ["Fare Tekerleği Kaydırma Hızı"] = "Velocidad de rueda del ratón",
                ["Kaydırma Hızı"] = "Velocidad de desplazamiento",
                ["Sayfalar arasında dikey kaydırma yaparken kullanılan hız çarpanı."] = "Multiplicador de velocidad al desplazarse verticalmente por las páginas.",
                ["Varsayılan Yakınlaştırma (Zoom)"] = "Nivel de zoom predeterminado",
                ["Varsayılan Yakınlaştırma"] = "Zoom predeterminado",
                ["Yeni bir PDF açıldığında uygulanacak başlangıç yakınlaştırma oranı."] = "Porcentaje de zoom initial aplicado al abrir un nuevo PDF.",
                ["Varsayılan Ek Çözüm Kağıdı"] = "Hoja de resolución predeterminada",
                ["Ek Çözüm Kağıdı"] = "Hoja lateral de resolución",
                ["PDF sayfalarının sağ tarafında ekstra işlem ve çözüm alanı açar."] = "Abre un área adicional de operaciones y resolución a la derecha del PDF.",
                ["Ek Çözüm Kağıdı Genişliği"] = "Ancho de hoja lateral",
                ["Sağ taraftaki ek çözüm kağıdının piksel cinsinden genişliği."] = "Ancho en píxeles de la hoja de resolución situada a la derecha.",
                ["Çözüm Kağıdı Deseni"] = "Patrón de hoja de resolución",
                ["Kağıt Deseni"] = "Patrón de papel",
                ["Ek çözüm kağıdında kullanılacak varsayılan kılavuz çizgi deseni."] = "Patrón de líneas guía predeterminado para la hoja de resolución.",
                ["Boş Beyaz"] = "Blanco liso",
                ["Kareli"] = "Cuadriculado",
                ["Çizgili"] = "Rayado",
                ["Noktalı"] = "Punteado",
                ["Kılavuz Çizgi / Izgara Rengi"] = "Color de línea guía / cuadrícula",
                ["Çizgi / Izgara Rengi"] = "Color de línea / cuadrícula",
                ["Kareli, çizgili veya noktalı kağıt deseninin çizgi rengi."] = "Color de línea para patrones cuadriculados, rayados o punteados.",
                ["Koyu Gri"] = "Gris oscuro",
                ["Açık Gri"] = "Gris claro",
                ["Klasik Mavi"] = "Azul clásico",
                ["Mavi"] = "Azul",
                ["İndigo"] = "Índigo",
                ["Toplam Kitap / PDF"] = "Total de Libros / PDF",
                ["Çözümlü Sayfa"] = "Páginas anotadas",
                ["Disk Kullanımı"] = "Uso de disco",
                ["Veri ve Kayıt Klasörü"] = "Carpeta de datos y almacenamiento",
                ["Kütüphane Klasörü"] = "Carpeta de biblioteca",
                ["Tüm çizimleriniz ve PDF verileriniz bu klasörde saklanır."] = "Todos sus dibujos y metadatos PDF se guardan en esta carpeta.",
                ["Klasörü Aç"] = "Abrir carpeta",
                ["Önbelleği Temizle"] = "Limpiar caché",
                ["Geçici render belleğini temizleyerek RAM kullanımını sıfırlar."] = "Libera memoria RAM limpiando las imágenes temporales en caché.",
                ["Tüm Çizimleri Sıfırla"] = "Restablecer todos los dibujos",
                ["Kayıtlı tüm sayfa çizimlerini kalıcı olarak siler."] = "Elimina permanentemente todos los dibujos y anotaciones guardados.",
                ["Arayüz Dili"] = "Idioma de la interfaz",
                ["Uygulama arayüzü ve menüler için tercih edilen dili seçin."] = "Seleccione el idioma preferido para la interfaz y los menús.",
                ["Başlangıçta Açılış Davranışı"] = "Comportamiento al iniciar",
                ["SarpNote başlatıldığında ilk olarak hangi ekranın açılacağını seçin."] = "Elija qué pantalla se abre primero al iniciar SarpNote.",
                ["Kitaplık Ekranını Aç"] = "Abrir biblioteca",
                ["Son Açılan PDF'den Devam Et"] = "Continuar último PDF abierto",
                ["Boş Çalışma Alanı Aç"] = "Abrir espacio de trabajo en blanco",
                ["Varsayılan Başlangıç Aracı"] = "Herramienta inicial predeterminada",
                ["Başlangıç Aracı"] = "Herramienta inicial",
                ["PDF açıldığında otomatik seçili gelecek çizim aracı."] = "Herramienta de dibujo seleccionada automáticamente al abrir un PDF.",
                ["Kalem"] = "Lápiz",
                ["Fosforlu Kalem"] = "Resaltador",
                ["Silgi"] = "Borrador",
                ["El (Kaydırma)"] = "Mano (Desplazar)",
                ["Başlangıç Kalem Kalınlığı"] = "Grosor inicial del lápiz",
                ["Uygulama açıldığında kalemin başlayacağı varsayılan uç kalınlığı."] = "Grosor predeterminado de la punta del lápiz al iniciar la aplicación.",
                ["Başlangıç Silgi Boyutu"] = "Tamaño inicial del borrador",
                ["Uygulama açıldığında silginin başlayacağı varsayılan boyut."] = "Tamaño predeterminado del borrador al iniciar la aplicación.",
                ["Başlangıç Pencere Modu"] = "Modo de ventana al iniciar",
                ["Pencere Durumu"] = "Estado de la ventana",
                ["Uygulamanın tam ekran mı yoksa normal pencerede mi açılacağını belirler."] = "Determina si la aplicación inicia maximizada o en ventana normal.",
                ["Tam Ekran (Maksimum)"] = "Maximizada (Pantalla completa)",
                ["Normal Pencere"] = "Ventana normal",
                ["Pen Studio"] = "Estudio de Lápiz",
                ["Kalem Stüdyosu"] = "Estudio de Lápiz",
                ["Kalınlık"] = "Grosor",
                ["Renkler"] = "Colores",
                ["Opaklık"] = "Opacidad",
                ["- - -  Kesikli Çizgi"] = "- - -  Línea discontinua",
                ["➔  Vektör Oku"] = "➔  Flecha vectorial",
                ["Hızlı Favori Kalemler (Sol Tık: Seç | Sağ Tık: Kaydet)"] = "Lápices favoritos (Clic izq.: Elegir | Clic der.: Guardar)",
                ["◫ Cevap Anahtarı / Bölünmüş Sayfa"] = "◫ Clave de respuestas / Vista dividida",
                ["Son"] = "Final",
                ["SarpNote - Kırpılan Soru"] = "SarpNote - Pregunta recortada",
                ["✓ Soru panoya kopyalandı (Ctrl+V ile yapıştırabilir veya aşağıdan kaydedebilirsin):"] = "✓ Pregunta copiada al portapapeles (Pegue con Ctrl+V o guarde abajo):",
                ["📌 Hata Defterine Kaydet"] = "📌 Guardar en cuaderno de errores",
                ["💾 Farklı Kaydet (PNG)..."] = "💾 Guardar como (PNG)..."
            }
        };

        private string GetSavedLanguageFilePath()
        {
            string dir = !string.IsNullOrEmpty(_baseDir)
                ? _baseDir
                : IOPath.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SarpNote_Library");
            Directory.CreateDirectory(dir);
            return IOPath.Combine(dir, "language.cfg");
        }

        private string LoadPersistedLanguageCode()
        {
            try
            {
                string f = GetSavedLanguageFilePath();
                if (File.Exists(f))
                {
                    string c = File.ReadAllText(f, Encoding.UTF8).Trim().ToLowerInvariant();
                    if (c == "tr" || _i18nDict.ContainsKey(c)) return c;
                }
            }
            catch { }
            return !string.IsNullOrWhiteSpace(_meta?.Language) ? _meta.Language.Trim().ToLowerInvariant() : "tr";
        }

        private void SavePersistedLanguageCode(string langCode)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(langCode)) return;
                string clean = langCode.Trim().ToLowerInvariant();
                File.WriteAllText(GetSavedLanguageFilePath(), clean, Encoding.UTF8);
                if (_meta != null) { _meta.Language = clean; SaveMetadata(); }
            }
            catch { }
        }

        private void InvokeNativeDashboardRefresh()
        {
            try
            {
                var methods = GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                foreach (var m in methods)
                {
                    if (m.GetParameters().Length == 0 && m.ReturnType == typeof(void))
                    {
                        string n = m.Name;
                        if (n.Equals("RefreshDashboard", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("RenderDashboard", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("UpdateDashboard", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("PopulateDashboard", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("RefreshLibraryUI", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("RenderLibrary", StringComparison.OrdinalIgnoreCase) ||
                            n.Equals("BuildDashboardList", StringComparison.OrdinalIgnoreCase))
                        {
                            m.Invoke(this, null);
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        

                private void AddRecoveredBookItem(System.Collections.IList bookList, Type itemType, string fullPath, int lastPageZeroBased, string subject)
        {
            try
            {
                object item = Activator.CreateInstance(itemType);
                string fileName = IOPath.GetFileName(fullPath);
                string dirName = IOPath.GetFileName(IOPath.GetDirectoryName(fullPath) ?? "PDFs");

                foreach (var p in itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (!p.CanWrite) continue;
                    string pn = p.Name;

                    if (p.PropertyType == typeof(string))
                    {
                        if (pn.Equals("Id", StringComparison.OrdinalIgnoreCase))
                            p.SetValue(item, Guid.NewGuid().ToString("N"));
                        else if (pn.IndexOf("Path", StringComparison.OrdinalIgnoreCase) >= 0)
                            p.SetValue(item, fullPath);
                        else if (pn.Equals("Title", StringComparison.OrdinalIgnoreCase) || pn.Equals("Name", StringComparison.OrdinalIgnoreCase) || pn.Equals("FileName", StringComparison.OrdinalIgnoreCase))
                            p.SetValue(item, fileName);
                        else if (pn.IndexOf("Subject", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Tag", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Course", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Category", StringComparison.OrdinalIgnoreCase) >= 0)
                            p.SetValue(item, subject);
                        else if (pn.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0)
                            p.SetValue(item, $"../{dirName}");
                    }
                    else if (p.PropertyType == typeof(int))
                    {
                        if (pn.IndexOf("LastPage", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("CurrentPage", StringComparison.OrdinalIgnoreCase) >= 0 || pn.Equals("Page", StringComparison.OrdinalIgnoreCase))
                            p.SetValue(item, lastPageZeroBased);
                    }
                    else if (p.PropertyType == typeof(DateTime))
                    {
                        p.SetValue(item, DateTime.Now);
                    }
                }

                bookList.Add(item);
            }
            catch { }
        }

        private string TrLoadingText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            string t = text.Trim();
            if (_currentLang == "tr") return text;

            var mFileLoad = Regex.Match(t, @"^(.+?)\s+(?:yükleniyor|hazırlanıyor|açılıyor|işleniyor)\.*$", RegexOptions.IgnoreCase);
            if (mFileLoad.Success)
            {
                string subj = mFileLoad.Groups[1].Value.Trim();
                if (subj.IndexOf("Kütüphane", StringComparison.OrdinalIgnoreCase) < 0 &&
                    subj.IndexOf("Kitaplık", StringComparison.OrdinalIgnoreCase) < 0 &&
                    subj.IndexOf("Sayfa", StringComparison.OrdinalIgnoreCase) < 0 &&
                    subj.IndexOf("Motor", StringComparison.OrdinalIgnoreCase) < 0 &&
                    subj.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return _currentLang switch
                    {
                        "en" => $"Loading {subj}...",
                        "de" => $"{subj} wird geladen...",
                        "fr" => $"Chargement de {subj}...",
                        "es" => $"Cargando {subj}...",
                        _    => t
                    };
                }
            }

            var mProg = Regex.Match(t, @"(\(\s*\d+\s*/\s*\d+\s*\)|\d+\s*%)$");
            string suffix = mProg.Success ? " " + mProg.Value : "";

            if (t.IndexOf("Kütüphane", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Kitaplık", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Workspace", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Çalışma alanı", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Preparing library and workspace...", "de" => "Bibliothek und Arbeitsbereich werden vorbereitet...", "fr" => "Préparation de la bibliothèque...", "es" => "Preparando biblioteca y espacio de trabajo...", _ => t }) + suffix;

            if (t.IndexOf("Motor", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("200Hz", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Donanım", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Kalem", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Initializing 200Hz drawing engine...", "de" => "200Hz-Zeichen-Engine wird gestartet...", "fr" => "Initialisation du moteur 200Hz...", "es" => "Iniciando motor de dibujo 200Hz...", _ => t }) + suffix;

            if (t.IndexOf("Sayfa", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Render", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Rendering high-resolution pages...", "de" => "Seiten werden gerendert...", "fr" => "Rendu des pages...", "es" => "Renderizando páginas...", _ => t }) + suffix;

            if (t.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Belge", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Kitap", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Loading PDF document...", "de" => "PDF-Dokument wird geladen...", "fr" => "Chargement du document PDF...", "es" => "Cargando documento PDF...", _ => t }) + suffix;

            if (t.IndexOf("Çizim", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Not", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Restoring saved annotations...", "de" => "Gespeicherte Zeichnungen werden geladen...", "fr" => "Restauration des annotations...", "es" => "Restaurando anotaciones...", _ => t }) + suffix;

            if (t.IndexOf("Hazır", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Tamam", StringComparison.OrdinalIgnoreCase) >= 0)
                return _currentLang switch { "en" => "Ready!", "de" => "Bereit!", "fr" => "Prêt !", "es" => "¡Listo!", _ => t };

            if (t.IndexOf("Yükleniyor", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Hazırlanıyor", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Başlatılıyor", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("bekleyin", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("SarpNote", StringComparison.OrdinalIgnoreCase) >= 0)
                return (_currentLang switch { "en" => "Starting SarpNote, please wait...", "de" => "SarpNote wird gestartet, bitte warten...", "fr" => "Démarrage de SarpNote...", "es" => "Iniciando SarpNote, espere...", _ => t }) + suffix;

            return TrText(t);
        }

        private string TrText(string trSource)
        {
            if (string.IsNullOrWhiteSpace(trSource)) return trSource;

            string trimmed = Regex.Replace(trSource.Trim(), @"\s+", " ");

            string dyn = TranslateDynamicPattern(trimmed);
            if (dyn != trimmed) return dyn;

            if (trimmed.Equals("+ Kitap Ekle", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("+ Add Book", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("+ Buch hinzufügen", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("+ Ajouter un livre", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("+ Añadir libro", StringComparison.OrdinalIgnoreCase))
            {
                return _currentLang switch
                {
                    "en" => "+ Add Book",
                    "de" => "+ Buch hinzufügen",
                    "fr" => "+ Ajouter un livre",
                    "es" => "+ Añadir libro",
                    _    => "+ Kitap Ekle"
                };
            }

            if (trimmed.Equals("Kitap Ekle", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("Add Book", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("Buch hinzufügen", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("Ajouter un livre", StringComparison.OrdinalIgnoreCase) ||
                trimmed.Equals("Añadir libro", StringComparison.OrdinalIgnoreCase))
            {
                return _currentLang switch
                {
                    "en" => "Add Book",
                    "de" => "Buch hinzufügen",
                    "fr" => "Ajouter un livre",
                    "es" => "Añadir libro",
                    _    => "Kitap Ekle"
                };
            }

            if (_currentLang == "tr")
            {
                if (_i18nDict.TryGetValue("tr", out var trMap) && trMap.TryGetValue(trimmed, out string trExact))
                    return trExact;
                return trSource;
            }

            if (_i18nDict.TryGetValue(_currentLang, out var map))
            {
                if (map.TryGetValue(trimmed, out string exact))
                    return exact;
                if (map.TryGetValue(trSource.Trim(), out string rawExact))
                    return rawExact;

                var prefixMatch = Regex.Match(trimmed, @"^([\+\-\•\★\☆\⏱\📐\🔬\🚀\🎯\◫\✂\📤\📌\💾\🔄\↺\🏷]+)\s*(.+)$");
                if (prefixMatch.Success)
                {
                    string bodyKey = prefixMatch.Groups[2].Value.Trim();
                    if (map.TryGetValue(bodyKey, out string bodyTrans))
                    {
                        return $"{prefixMatch.Groups[1].Value} {bodyTrans}";
                    }
                }
            }

            return trSource;
        }

        private string TranslateDynamicPattern(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            var mParenPages = Regex.Match(
                text,
                @"^(?:Sayfalar|Sayfa|Pages|Page|Seiten|Seite|Páginas|Página)\s*\(\s*(\d+)\s*\)$",
                RegexOptions.IgnoreCase);
            if (mParenPages.Success)
            {
                string num = mParenPages.Groups[1].Value;
                string pWord = _currentLang switch
                {
                    "en" => "Pages",
                    "de" => "Seiten",
                    "fr" => "Pages",
                    "es" => "Páginas",
                    _    => "Sayfalar"
                };
                return $"{pWord} ({num})";
            }

            if (Regex.IsMatch(text, @"^(?:Sayfalar|Pages|Seiten|Páginas)$", RegexOptions.IgnoreCase))
            {
                return _currentLang switch
                {
                    "en" => "Pages",
                    "de" => "Seiten",
                    "fr" => "Pages",
                    "es" => "Páginas",
                    _    => "Sayfalar"
                };
            }

            var mPageOf = Regex.Match(text, @"^(?:Sayfa|Page|Seite|Página)\s*(\d+)\s*/\s*(\d+)$", RegexOptions.IgnoreCase);
            if (mPageOf.Success)
            {
                string pWord = _currentLang switch { "en" => "Page", "de" => "Seite", "fr" => "Page", "es" => "Página", _ => "Sayfa" };
                return $"{pWord} {mPageOf.Groups[1].Value} / {mPageOf.Groups[2].Value}";
            }

            var mSinglePage = Regex.Match(text, @"^(?:Sayfa|Page|Seite|Página)\s*(\d+)$", RegexOptions.IgnoreCase);
            if (mSinglePage.Success)
            {
                string pWord = _currentLang switch { "en" => "Page", "de" => "Seite", "fr" => "Page", "es" => "Página", _ => "Sayfa" };
                return $"{pWord} {mSinglePage.Groups[1].Value}";
            }

            var mPagesCount = Regex.Match(text, @"^(\d+)\s*(?:Sayfa|sayfa|Pages|Seiten|Páginas)$", RegexOptions.IgnoreCase);
            if (mPagesCount.Success)
            {
                string pWord = _currentLang switch { "en" => "Pages", "de" => "Seiten", "fr" => "Pages", "es" => "Páginas", _ => "Sayfa" };
                return $"{mPagesCount.Groups[1].Value} {pWord}";
            }

            var mBooksCount = Regex.Match(text, @"^(\d+)\s*(?:Kitap|Belge|Books|Documents|Bücher|Livres|Libros)$", RegexOptions.IgnoreCase);
            if (mBooksCount.Success)
            {
                string bWord = _currentLang switch { "en" => "Books", "de" => "Bücher", "fr" => "Livres", "es" => "Libros", _ => "Kitap" };
                return $"{mBooksCount.Groups[1].Value} {bWord}";
            }

            var mTimeAgo = Regex.Match(text, @"^(\d+)\s*(saat|dk|dakika|gün|hours?|mins?|minutes?|days?|Std\.?|Min\.?|Tagen?|h|jours?|horas?|días?)\s*(?:önce|ago|vor|il y a|hace)?$", RegexOptions.IgnoreCase);
            if (mTimeAgo.Success)
            {
                string num = mTimeAgo.Groups[1].Value;
                string unit = mTimeAgo.Groups[2].Value.ToLowerInvariant();
                bool isHour = unit.StartsWith("saat") || unit.StartsWith("hour") || unit.StartsWith("std") || unit == "h" || unit.StartsWith("hora");
                bool isDay  = unit.StartsWith("gün") || unit.StartsWith("day") || unit.StartsWith("tag") || unit.StartsWith("jour") || unit.StartsWith("día");

                return _currentLang switch
                {
                    "en" => isDay ? $"{num} days ago" : (isHour ? $"{num} hours ago" : $"{num} mins ago"),
                    "de" => isDay ? $"vor {num} Tagen" : (isHour ? $"vor {num} Std." : $"vor {num} Min."),
                    "fr" => isDay ? $"il y a {num} j" : (isHour ? $"il y a {num} h" : $"il y a {num} min"),
                    "es" => isDay ? $"hace {num} días" : (isHour ? $"hace {num} h" : $"hace {num} min"),
                    _    => isDay ? $"{num} gün önce" : (isHour ? $"{num} saat önce" : $"{num} dk önce")
                };
            }

            var mLastPage = Regex.Match(text, @"^(?:Son Sayfa|Kaldığın Sayfa|Last Page|Letzte Seite|Dernière page|Última página)\s*:\s*(\d+)$", RegexOptions.IgnoreCase);
            if (mLastPage.Success)
            {
                string lpWord = _currentLang switch { "en" => "Last Page", "de" => "Letzte Seite", "fr" => "Dernière page", "es" => "Última página", _ => "Son Sayfa" };
                return $"{lpWord}: {mLastPage.Groups[1].Value}";
            }

            return text;
        }

        private void EnsureLanguageSelectorInSettings()
        {
            AttachLocalizationHooksOnce();

            if (_langSettingCard != null || SetPageStartup == null) return;
            Panel targetHost = SetPageStartup;

            _langSettingCard = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(22, 26, 38)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(42, 48, 68)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(16, 12, 16, 12),
                Margin = new Thickness(0, 0, 0, 10)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var leftStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 18, 0) };
            var titleBlock = new TextBlock
            {
                Text = "Arayüz Dili",
                Foreground = Brushes.WhiteSmoke,
                FontSize = 13.5,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap
            };
            var descBlock = new TextBlock
            {
                Text = "Uygulama arayüzü ve menüler için tercih edilen dili seçin.",
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                FontSize = 11.5,
                Margin = new Thickness(0, 3, 0, 0),
                TextWrapping = TextWrapping.Wrap
            };
            leftStack.Children.Add(titleBlock);
            leftStack.Children.Add(descBlock);
            Grid.SetColumn(leftStack, 0);

            _cmbAppLanguage = new ComboBox
            {
                Width = 235,
                Height = 34,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 12.5,
                FontWeight = FontWeights.SemiBold
            };
            if (TryFindResource("SettingsComboBox") is Style cmbStyle)
                _cmbAppLanguage.Style = cmbStyle;

            var langs = new[]
            {
                ("tr", "🇹🇷  Türkçe"),
                ("en", "🇬🇧  English"),
                ("de", "🇩🇪  Deutsch"),
                ("fr", "🇫🇷  Français"),
                ("es", "🇪🇸  Español")
            };
            foreach (var (code, label) in langs)
            {
                _cmbAppLanguage.Items.Add(new ComboBoxItem
                {
                    Tag = code,
                    Content = label,
                    Padding = new Thickness(10, 6, 10, 6)
                });
            }

            _cmbAppLanguage.SelectionChanged += (s, e) =>
            {
                if (_isApplyingLanguage) return;
                if (_cmbAppLanguage.SelectedItem is ComboBoxItem item && item.Tag is string langCode)
                {
                    string target = langCode.ToLowerInvariant();
                    _currentLang = target;
                    SavePersistedLanguageCode(target);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ApplyApplicationLanguage(target, saveToMeta: true);
                    }), System.Windows.Threading.DispatcherPriority.Background);
                }
            };

            Grid.SetColumn(_cmbAppLanguage, 1);
            grid.Children.Add(leftStack);
            grid.Children.Add(_cmbAppLanguage);
            _langSettingCard.Child = grid;

            targetHost.Children.Insert(0, _langSettingCard);
            SyncLanguageComboSelection();
        }

                private void AttachLocalizationHooksOnce()
        {
            if (_localizationHooksAttached) return;
            _localizationHooksAttached = true;

            // Sadece Kitaplık veya Ayarlar alt sayfalarının görünürlüğü değiştiğinde çeviriyi çalıştır (Tuş tıklamalarına müdahale etmez!)
            if (DashboardOverlay != null)
            {
                DashboardOverlay.IsVisibleChanged += (s, e) =>
                {
                    if (DashboardOverlay.IsVisible && !_isApplyingLanguage)
                        ForceTranslateDashboardAndPagesHeader();
                };
            }

            var settingsPages = new FrameworkElement[] { SetPagePen, SetPagePerf, SetPageWork, SetPageStorage, SetPageStartup };
            foreach (var page in settingsPages)
            {
                if (page != null)
                {
                    page.IsVisibleChanged += (s, e) =>
                    {
                        if (page.IsVisible && !_isApplyingLanguage)
                            ApplyStructuralSettingsTranslations();
                    };
                }
            }
        }

        private void SyncLanguageComboSelection()
        {
            if (_cmbAppLanguage == null) return;
            bool prevApplying = _isApplyingLanguage;
            _isApplyingLanguage = true;
            try
            {
                ComboBoxItem matched = null;
                for (int i = 0; i < _cmbAppLanguage.Items.Count; i++)
                {
                    if (_cmbAppLanguage.Items[i] is ComboBoxItem item &&
                        string.Equals(item.Tag as string, _currentLang, StringComparison.OrdinalIgnoreCase))
                    {
                        matched = item;
                        if (_cmbAppLanguage.SelectedIndex != i)
                            _cmbAppLanguage.SelectedIndex = i;
                        break;
                    }
                }

                if (matched?.Content is string labelStr)
                {
                    _cmbAppLanguage.ApplyTemplate();
                    _cmbAppLanguage.UpdateLayout();
                    foreach (var tb in FindAllChildren<TextBlock>(_cmbAppLanguage))
                    {
                        if (!FindAllChildren<ComboBoxItem>(_cmbAppLanguage).Any(ci => tb.IsDescendantOf(ci)))
                        {
                            tb.Text = labelStr;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                _isApplyingLanguage = prevApplying;
            }
        }

        private void ApplyApplicationLanguage(string langCode, bool saveToMeta)
        {
            if (_isApplyingLanguage) return;
            _isApplyingLanguage = true;

            try
            {
                if (string.IsNullOrWhiteSpace(langCode)) langCode = "tr";
                langCode = langCode.ToLowerInvariant();
                if (!_i18nDict.ContainsKey(langCode))
                    langCode = "tr";

                _currentLang = langCode;
                if (_meta != null)
                {
                    _meta.Language = _currentLang;
                    if (saveToMeta) SavePersistedLanguageCode(_currentLang);
                }

                SyncLanguageComboSelection();

                _structurallyTranslatedElements.Clear();

                var visited = new HashSet<DependencyObject>();
                TranslateElementTree(this, visited);

                ApplyStructuralSettingsTranslations();
                ForceTranslateDashboardAndPagesHeader();
            }
            catch { }
            finally
            {
                _isApplyingLanguage = false;
            }
        }

        private void ForceTranslateDashboardAndPagesHeader()
        {
            try
            {
                foreach (var tb in FindAllChildren<TextBlock>(this))
                {
                    string name = tb.Name ?? "";
                    string txt = (tb.Text ?? "").Trim();
                    if (string.IsNullOrEmpty(txt)) continue;

                    if (name.IndexOf("Load", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("Splash", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("Progress", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("Wait", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        tb.Text = TrLoadingText(txt);
                    }

                    // Tek başına "Pages" veya "Pages (176)" başlığı
                    var mPages = Regex.Match(
                        txt,
                        @"^(?:Sayfalar|Sayfa|Pages|Page|Seiten|Seite|Páginas|Página)(?:\s*\(\s*(\d+)\s*\))?$",
                        RegexOptions.IgnoreCase);
                    if (mPages.Success && ( mPages.Groups[1].Success || txt.Equals("Pages", StringComparison.OrdinalIgnoreCase) || txt.Equals("Sayfalar", StringComparison.OrdinalIgnoreCase) || txt.Equals("Seiten", StringComparison.OrdinalIgnoreCase) || txt.Equals("Páginas", StringComparison.OrdinalIgnoreCase)))
                    {
                        string pWord = _currentLang switch
                        {
                            "en" => "Pages",
                            "de" => "Seiten",
                            "fr" => "Pages",
                            "es" => "Páginas",
                            _    => "Sayfalar"
                        };
                        tb.Text = mPages.Groups[1].Success ? $"{pWord} ({mPages.Groups[1].Value})" : pWord;
                        _structurallyTranslatedElements.Add(tb);
                        continue;
                    }

                    if (txt.Equals("+ Kitap Ekle", StringComparison.OrdinalIgnoreCase) ||
                        txt.Equals("+ Add Book", StringComparison.OrdinalIgnoreCase) ||
                        txt.Equals("+ Buch hinzufügen", StringComparison.OrdinalIgnoreCase) ||
                        txt.Equals("+ Ajouter un livre", StringComparison.OrdinalIgnoreCase) ||
                        txt.Equals("+ Añadir libro", StringComparison.OrdinalIgnoreCase))
                    {
                        tb.Text = _currentLang switch
                        {
                            "en" => "+ Add Book",
                            "de" => "+ Buch hinzufügen",
                            "fr" => "+ Ajouter un livre",
                            "es" => "+ Añadir libro",
                            _    => "+ Kitap Ekle"
                        };
                        _structurallyTranslatedElements.Add(tb);
                        continue;
                    }
                }

                if (DashboardOverlay != null)
                {
                    foreach (var btn in FindAllChildren<Button>(DashboardOverlay))
                    {
                        if (btn.Content is string sContent)
                        {
                            string t = sContent.Trim();
                            if (t.IndexOf("Kitap Ekle", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                t.IndexOf("Add Book", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                t.IndexOf("Buch hinzufügen", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                t.IndexOf("Ajouter un livre", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                t.IndexOf("Añadir libro", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                btn.Content = _currentLang switch
                                {
                                    "en" => "+ Add Book",
                                    "de" => "+ Buch hinzufügen",
                                    "fr" => "+ Ajouter un livre",
                                    "es" => "+ Añadir libro",
                                    _    => "+ Kitap Ekle"
                                };
                                _structurallyTranslatedElements.Add(btn);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void ApplyStructuralSettingsTranslations()
        {
            if (SettingsOverlay == null) return;

            foreach (var btn in FindAllChildren<Button>(SettingsOverlay))
            {
                string tagOrName = $"{btn.Tag} {btn.Name}";
                if (btn == BtnSetPenBasic || btn == BtnSetPenPro || btn == BtnQualFast || btn == BtnQualBalanced || btn == BtnQualUltra || btn == BtnFilterNormal || btn == BtnFilterSepia || btn == BtnFilterDim)
                    continue;

                if (tagOrName.IndexOf("Pen", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Kalem & Silgi");
                else if (tagOrName.IndexOf("Perf", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Görüntü & Performans");
                else if (tagOrName.IndexOf("Work", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Çalışma Alanı");
                else if (tagOrName.IndexOf("Storage", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Depolama");
                else if (tagOrName.IndexOf("Startup", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Başlangıç");
                else if (tagOrName.IndexOf("Reset", StringComparison.OrdinalIgnoreCase) >= 0)
                    TranslateButtonPrimaryText(btn, "Varsayılana Sıfırla");
                else if (tagOrName.IndexOf("Close", StringComparison.OrdinalIgnoreCase) >= 0 || tagOrName.IndexOf("Save", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var tbs = FindAllChildren<TextBlock>(btn);
                    if (tbs.Count > 0 && tbs[0].Text.Length > 2)
                        TranslateButtonPrimaryText(btn, "Kaydet ve Kapat");
                }
            }

            if (TxtSettingsPageTitle != null && TxtSettingsPageSub != null)
            {
                if (SetPagePen?.Visibility == Visibility.Visible)
                {
                    SetTranslatedTextBlock(TxtSettingsPageTitle, "Kalem ve Silgi Motoru");
                    SetTranslatedTextBlock(TxtSettingsPageSub, "Çizim hassasiyeti, şekil tanıma ve silgi davranışlarını özelleştirin.");
                }
                else if (SetPagePerf?.Visibility == Visibility.Visible)
                {
                    SetTranslatedTextBlock(TxtSettingsPageTitle, "Görüntü & Performans");
                    SetTranslatedTextBlock(TxtSettingsPageSub, "4K PDF render kalitesi, okuma filtreleri ve kaydırma akıcılığını ayarlayın.");
                }
                else if (SetPageWork?.Visibility == Visibility.Visible)
                {
                    SetTranslatedTextBlock(TxtSettingsPageTitle, "Çalışma Alanı & Kağıt");
                    SetTranslatedTextBlock(TxtSettingsPageSub, "Varsayılan yakınlaştırma, ek çözüm kağıdı ve kılavuz çizgi tercihlerini yönetin.");
                }
                else if (SetPageStorage?.Visibility == Visibility.Visible)
                {
                    SetTranslatedTextBlock(TxtSettingsPageTitle, "Depolama ve Veri Yönetimi");
                    SetTranslatedTextBlock(TxtSettingsPageSub, "Kütüphane istatistiklerini görüntüleyin, önbelleği ve kayıt klasörünü yönetin.");
                }
                else if (SetPageStartup?.Visibility == Visibility.Visible)
                {
                    SetTranslatedTextBlock(TxtSettingsPageTitle, "Başlangıç Ayarları");
                    SetTranslatedTextBlock(TxtSettingsPageSub, "Uygulama açılışındaki varsayılan ekran, araç ve pencere tercihlerini belirleyin.");
                }
            }

            TranslateModeButtonPair(BtnSetPenBasic, "Basic (Sabit Uç)", "Basınçsız, sabit kalınlık");
            TranslateModeButtonPair(BtnSetPenPro, "Pro (Basınç Duyarlı)", "Basınç destekli dinamik uç");
            TranslateModeButtonPair(BtnQualFast, "Hızlı", "Düşük RAM");
            TranslateModeButtonPair(BtnQualBalanced, "Dengeli", "Önerilen");
            TranslateModeButtonPair(BtnQualUltra, "Ultra (4K)", "Maksimum Netlik");
            TranslateModeButtonPair(BtnFilterNormal, "Normal", "Orijinal Beyaz");
            TranslateModeButtonPair(BtnFilterSepia, "Sepya", "Sıcak Ton");
            TranslateModeButtonPair(BtnFilterDim, "Gece / Koyu", "Loş");

            TranslateCardByAnchorControl(BtnSetPenBasic, "Varsayılan Kalem Modu", "Çizim yaparken kullanılacak varsayılan kalem motorunu seçin.");
            TranslateCardByAnchorControl(TxtSetHoldMs, "Basılı Tutma Süresi (Otomatik Şekil)", "Çizgiyi düzleştirmek veya şekle dönüştürmek için gereken bekleme süresi.");
            TranslateCardByAnchorControl(TxtSetSmoothing, "Çizgi Yumuşatma (Smoothing)", "El titremesini azaltır ve çizgilerin daha pürüzsüz görünmesini sağlar.");

            TranslateCardByAnchorControl(BtnQualFast, "PDF Render Kalitesi", "Sayfa netliği ve bellek kullanımı arasındaki dengeyi belirler.");
            TranslateCardByAnchorControl(BtnFilterNormal, "Sayfa Okuma Filtresi", "Uzun süreli çalışmalarda göz yorgunluğunu azaltmak için sayfa tonunu ayarlar.");
            TranslateCardByAnchorControl(TxtSetCachePages, "Önbellekte Tutulacak Sayfa Sayısı", "Aktif sayfanın çevresinde yüksek çözünürlükte hazır tutulan sayfa sayısı.");
            TranslateCardByAnchorControl(TxtSetScrollSpeed, "Fare Tekerleği Kaydırma Hızı", "Sayfalar arasında dikey kaydırma yaparken kullanılan hız çarpanı.");

            TranslateCardByAnchorControl(TxtSetDefZoom, "Varsayılan Yakınlaştırma (Zoom)", "Yeni bir PDF açıldığında uygulanacak başlangıç yakınlaştırma oranı.");
            TranslateCardByAnchorControl(TxtSetMarginWidth, "Ek Çözüm Kağıdı Genişliği", "Sağ taraftaki ek çözüm kağıdının piksel cinsinden genişliği.");

            TranslateCardByAnchorControl(_cmbAppLanguage, "Arayüz Dili", "Uygulama arayüzü ve menüler için tercih edilen dili seçin.");
            TranslateCardByAnchorControl(SliderSetStartPenSize, "Başlangıç Kalem Kalınlığı", "Uygulama açıldığında kalemin başlayacağı varsayılan uç kalınlığı.");
            TranslateCardByAnchorControl(SliderSetStartEraserSize, "Başlangıç Silgi Boyutu", "Uygulama açıldığında silginin başlayacağı varsayılan boyut.");

            TranslateStoragePageElements();

            TranslatePageCardsByControlType(SetPagePen,
                checkCards: new[] { ("Basılı Tutunca Şekle Dönüştür", "Çizgi veya kapalı şekil çizip ucunda bekletince düzgün geometriye dönüştürür.") },
                comboCards: new[] { ("Varsayılan Silgi Modu", "Silginin çizgileri tek dokunuşla mı yoksa bölgesel mi sileceğini belirler.", new[] { "Çizgi Silgisi (Tüm Çizgiyi Siler)", "Nokta Silgisi (Kısmi Siler)" }) });

            TranslatePageCardsByControlType(SetPageWork,
                checkCards: new[] { ("Varsayılan Ek Çözüm Kağıdı", "PDF sayfalarının sağ tarafında ekstra işlem ve çözüm alanı açar.") },
                comboCards: new[]
                {
                    ("Çözüm Kağıdı Deseni", "Ek çözüm kağıdında kullanılacak varsayılan kılavuz çizgi deseni.", new[] { "Kareli", "Çizgili", "Noktalı", "Boş Beyaz" }),
                    ("Kılavuz Çizgi / Izgara Rengi", "Kareli, çizgili veya noktalı kağıt deseninin çizgi rengi.", new[] { "Koyu Gri", "Açık Gri", "Klasik Mavi", "İndigo" })
                });

            TranslatePageCardsByControlType(SetPageStartup,
                checkCards: new[] { ("Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran", "Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar.") },
                comboCards: new[]
                {
                    ("Başlangıçta Açılış Davranışı", "SarpNote başlatıldığında ilk olarak hangi ekranın açılacağını seçin.", new[] { "Kitaplık Ekranını Aç", "Son Açılan PDF'den Devam Et", "Boş Çalışma Alanı Aç" }),
                    ("Varsayılan Başlangıç Aracı", "PDF açıldığında otomatik seçili gelecek çizim aracı.", new[] { "Kalem", "Fosforlu Kalem", "Silgi", "El (Kaydırma)" }),
                    ("Başlangıç Pencere Modu", "Uygulamanın tam ekran mı yoksa normal pencerede mi açılacağını belirler.", new[] { "Tam Ekran (Maksimum)", "Normal Pencere" })
                });
        }

        private void TranslateStoragePageElements()
        {
            if (SetPageStorage == null) return;

            TranslateStatBoxLabel(TxtStatBooks, "Toplam Kitap / PDF");
            TranslateStatBoxLabel(TxtStatAnnotatedPages, "Çözümlü Sayfa");
            TranslateStatBoxLabel(TxtStatDiskSize, "Disk Kullanımı");

            var actionSpecs = new[]
            {
                ("Tüm Çözümleri ve Yıldızlı Soruları Yedekle", "Tüm el yazısı çözümlerini (.isf) ve metadata verilerini Masaüstüne tarihli klasör olarak kopyalar.", "📦 Masaüstüne Yedekle"),
                ("RAM ve Kapak Önbelleğini Ferahlat", "Hiçbir çizimi silmeden arka plandaki sayfa dokularını temizleyip belleği boşaltır.", "🧹 Önbelleği Temizle"),
                ("SarpNote Veri Klasörünü Aç", "Belgelerim/SarpNote_Library dizinini Windows Gezgini'nde görüntüler.", "📂 Klasörü Aç")
            };

            int cardIdx = 0;
            foreach (var border in SetPageStorage.Children.OfType<Border>())
            {
                var btns = FindAllChildren<Button>(border);
                if (btns.Count > 0 && cardIdx < actionSpecs.Length)
                {
                    var spec = actionSpecs[cardIdx++];
                    var headerTbs = GetCardHeaderTextBlocks(border);
                    if (headerTbs.Count >= 1) SetTranslatedTextBlock(headerTbs[0], spec.Item1);
                    if (headerTbs.Count >= 2 && !headerTbs[1].Text.Contains(":\\"))
                        SetTranslatedTextBlock(headerTbs[1], spec.Item2);
                    TranslateButtonPrimaryText(btns[0], spec.Item3);
                }
            }
        }

        private void TranslateStatBoxLabel(TextBlock statValueTb, string canonicalTrLabel)
        {
            if (statValueTb?.Parent is Panel p)
            {
                foreach (var tb in p.Children.OfType<TextBlock>())
                {
                    if (!ReferenceEquals(tb, statValueTb))
                    {
                        SetTranslatedTextBlock(tb, canonicalTrLabel);
                    }
                }
            }
        }

        private void TranslateButtonPrimaryText(Button btn, string canonicalTrText)
        {
            if (btn == null) return;
            if (btn.Content is string)
            {
                _structurallyTranslatedElements.Add(btn);
                btn.Content = TrText(canonicalTrText);
                return;
            }

            var tbs = FindAllChildren<TextBlock>(btn);
            if (tbs.Count == 1)
            {
                SetTranslatedTextBlock(tbs[0], canonicalTrText);
            }
            else if (tbs.Count >= 2)
            {
                var textTb = tbs.FirstOrDefault(t => Regex.IsMatch(t.Text ?? "", @"\p{L}{2,}")) ?? tbs[tbs.Count - 1];
                SetTranslatedTextBlock(textTb, canonicalTrText);
            }
        }

                private void TranslateModeButtonPair(Button btn, string trTitle, string trSub)
        {
            if (btn == null) return;
            btn.IsHitTestVisible = true;
            btn.Cursor = Cursors.Hand;

            // Butonun içindeki StackPanel/Grid arka planı null ise Transparent yap ki boşluklara tıklanınca da buton algılasın
            foreach (var panel in FindAllChildren<Panel>(btn))
            {
                panel.IsHitTestVisible = true;
                if (panel.Background == null)
                    panel.Background = Brushes.Transparent;
            }

            var tbs = FindAllChildren<TextBlock>(btn);
            foreach (var tb in tbs) tb.IsHitTestVisible = true;
            if (tbs.Count >= 1) SetTranslatedTextBlock(tbs[0], trTitle);
            if (tbs.Count >= 2) SetTranslatedTextBlock(tbs[1], trSub);
        }

        private void TranslateCardByAnchorControl(DependencyObject anchor, string trTitle, string trDesc)
        {
            if (anchor == null) return;
            Border cardBorder = FindEnclosingSettingCard(anchor);
            if (cardBorder == null) return;

            var headerTbs = GetCardHeaderTextBlocks(cardBorder);
            if (headerTbs.Count >= 1) SetTranslatedTextBlock(headerTbs[0], trTitle);
            if (headerTbs.Count >= 2) SetTranslatedTextBlock(headerTbs[1], trDesc);
        }

                private void TranslatePageCardsByControlType(
            FrameworkElement pageHost,
            (string Title, string Desc)[] checkCards,
            (string Title, string Desc, string[] ItemLabels)[] comboCards)
        {
            TranslateCardByAnchorControl(CmbSetBarrelMode, "Kalem Yan Tuşu (Barrel Button) Davranışı", "Tuşa basılı tutarken geçici silgi açabilir veya tek basışla kalıcı geçiş yapabilirsin.");
            TranslateCardByAnchorControl(ChkSetEraserAura, "Drawboard Silgi İzi ve Saydamlaştırma (Aura)", "Silgiyle dokunulan çizgileri silmeden önce %25 saydamlaştırır ve dış kontur gösterir.");
            TranslateCardByAnchorControl(ChkSetStartDrawer, "Başlangıçta Sol Sayfa Paneli Açık Olsun", "Uygulama açıldığında sol taraftaki küçük sayfa ön izleme çekmecesini otomatik açar.");
            TranslateCardByAnchorControl(CmbSetStartupAction, "Uygulama Açılış Davranışı", "SarpNote açıldığında son çalıştığın testi yükleyebilir veya Kitaplık ekranını açabilir.");
            TranslateCardByAnchorControl(CmbSetStartupPenStyle, "Başlangıç Varsayılan Kalem Türü", "Uygulama ilk açıldığında seçili gelecek kalem çizim modu (Pro veya Basic).");
            TranslateCardByAnchorControl(CmbSetStartupColor, "Başlangıç Varsayılan Kalem Rengi", "Açılışta varsayılan olarak seçili gelecek kalem mürekkep rengi.");
            TranslateCardByAnchorControl(ChkSetStartExtraMargin, "Yeni PDF'lerde Sağ Çözüm Kağıdı & Tam Ekran", "Açılan her testte sağ taraftaki işlem kağıdını otomatik hazırlar.");
            }

        private Border FindEnclosingSettingCard(DependencyObject control)
        {
            DependencyObject cur = control;
            while (cur != null)
            {
                DependencyObject parent = LogicalTreeHelper.GetParent(cur) ?? (cur is Visual ? VisualTreeHelper.GetParent(cur) : null);
                if (cur is Border bd && (ReferenceEquals(parent, SetPagePen) || ReferenceEquals(parent, SetPagePerf) || ReferenceEquals(parent, SetPageWork) || ReferenceEquals(parent, SetPageStorage) || ReferenceEquals(parent, SetPageStartup)))
                {
                    return bd;
                }
                cur = parent;
            }
            return null;
        }

        private List<TextBlock> GetCardHeaderTextBlocks(Border cardBorder)
        {
            if (cardBorder == null) return new List<TextBlock>();
            var valueBadges = new HashSet<TextBlock>
            {
                TxtSetHoldMs, TxtSetSmoothing, TxtSetCachePages, TxtSetScrollSpeed,
                TxtSetDefZoom, TxtSetMarginWidth, TxtSetStartPenSize, TxtSetStartEraserSize,
                TxtStatBooks, TxtStatAnnotatedPages, TxtStatDiskSize
            };

            return FindAllChildren<TextBlock>(cardBorder)
                .Where(tb => tb != null
                    && !valueBadges.Contains(tb)
                    && FindAncestor<Button>(tb) == null
                    && FindAncestor<ComboBox>(tb) == null
                    && FindAncestor<CheckBox>(tb) == null)
                .ToList();
        }

        private static List<T> FindAllChildren<T>(DependencyObject parent) where T : DependencyObject
        {
            var list = new List<T>();
            if (parent == null) return list;
            var visited = new HashSet<DependencyObject>();
            CollectChildrenRecursive(parent, list, visited);
            return list;
        }

        private static void CollectChildrenRecursive<T>(DependencyObject root, List<T> list, HashSet<DependencyObject> visited) where T : DependencyObject
        {
            if (root == null || !visited.Add(root)) return;
            if (root is T match) list.Add(match);

            foreach (var lChild in LogicalTreeHelper.GetChildren(root).OfType<DependencyObject>())
                CollectChildrenRecursive(lChild, list, visited);

            if (root is Visual || root is System.Windows.Media.Media3D.Visual3D)
            {
                int count = VisualTreeHelper.GetChildrenCount(root);
                for (int i = 0; i < count; i++)
                {
                    var vChild = VisualTreeHelper.GetChild(root, i);
                    if (vChild != null) CollectChildrenRecursive(vChild, list, visited);
                }
            }
        }

        private static T FindAncestor<T>(DependencyObject obj) where T : DependencyObject
        {
            DependencyObject cur = obj;
            while (cur != null)
            {
                if (cur is T target) return target;
                cur = LogicalTreeHelper.GetParent(cur) ?? (cur is Visual ? VisualTreeHelper.GetParent(cur) : null);
            }
            return null;
        }

        private void SetTranslatedTextBlock(TextBlock tb, string canonicalTrText)
        {
            if (tb == null) return;
            _structurallyTranslatedElements.Add(tb);
            _originalTextMap[tb] = canonicalTrText;

            string outText = _currentLang == "tr" ? canonicalTrText : TrText(canonicalTrText);
            tb.Text = outText;
            _lastAppliedTextMap[tb] = outText;
        }

        private void SetTranslatedComboBoxItem(ComboBoxItem cbi, string canonicalTrText)
        {
            if (cbi == null) return;
            _structurallyTranslatedElements.Add(cbi);

            string tag = (cbi.Tag as string)?.Trim() ?? "";
            string resolvedTr = tag switch
            {
                "LastPdf" or "Resume" or "Last" => "Son Açılan PDF'den Devam Et",
                "Library" or "Dashboard"        => "Kitaplık Ekranını Aç",
                "Blank" or "Empty" or "None"    => "Boş Çalışma Alanı Aç",
                "Pen"                           => "Kalem",
                "Highlighter"                   => "Fosforlu Kalem",
                "Eraser"                        => "Silgi",
                "Hand" or "Pan"                 => "El (Kaydırma)",
                "Maximized" or "Fullscreen"     => "Tam Ekran (Maksimum)",
                "Normal" or "Windowed"          => "Normal Pencere",
                "Stroke"                        => "Çizgi Silgisi (Tüm Çizgiyi Siler)",
                "Point" or "Partial"            => "Nokta Silgisi (Kısmi Siler)",
                "Grid" or "Squared"             => "Kareli",
                "Ruled" or "Lined"              => "Çizgili",
                "Dotted"                        => "Noktalı",
                "White" or "Plain"              => "Boş Beyaz",
                _                               => canonicalTrText
            };

            _originalTextMap[cbi] = resolvedTr;
            string outText = _currentLang == "tr" ? resolvedTr : TrText(resolvedTr);
            cbi.Content = outText;
            _lastAppliedTextMap[cbi] = outText;
        }

                private void RefreshComboBoxSelectionDisplay(ComboBox cmb)
        {
            if (cmb == null || ReferenceEquals(cmb, _cmbAppLanguage) || cmb.IsDropDownOpen) return;
            try
            {
                if (cmb.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content is string targetStr)
                {
                    foreach (var tb in FindAllChildren<TextBlock>(cmb))
                    {
                        if (!FindAllChildren<ComboBoxItem>(cmb).Any(item => tb.IsDescendantOf(item)))
                        {
                            _structurallyTranslatedElements.Add(tb);
                            if (tb.Text != targetStr) tb.Text = targetStr;
                        }
                    }
                }
            }
            catch { }
        }

        private void TranslateElementTree(DependencyObject root, HashSet<DependencyObject> visited)
        {
            if (root == null || visited == null || !visited.Add(root)) return;

            try
            {
                if (root is TextBlock tb && !_structurallyTranslatedElements.Contains(tb))
                {
                    if (tb.Inlines != null && tb.Inlines.Count > 1)
                    {
                        foreach (var inline in tb.Inlines)
                        {
                            if (inline is System.Windows.Documents.Run run && !string.IsNullOrWhiteSpace(run.Text))
                            {
                                string rCur = run.Text;
                                if (!_originalTextMap.TryGetValue(run, out string rOrig) ||
                                    (_lastAppliedTextMap.TryGetValue(run, out string rLast) && rCur != rLast))
                                {
                                    rOrig = rCur;
                                    _originalTextMap[run] = rOrig;
                                }
                                string leading = Regex.Match(rOrig, @"^\s*").Value;
                                string trailing = Regex.Match(rOrig, @"\s*$").Value;
                                string rTarget = leading + TrText(rOrig.Trim()) + trailing;
                                if (run.Text != rTarget) run.Text = rTarget;
                                _lastAppliedTextMap[run] = rTarget;
                            }
                        }
                    }
                    else
                    {
                        string curText = tb.Text ?? "";
                        if (!string.IsNullOrWhiteSpace(curText))
                        {
                            if (!_originalTextMap.TryGetValue(tb, out string orig) ||
                                (_lastAppliedTextMap.TryGetValue(tb, out string lastApp) && curText != lastApp))
                            {
                                orig = curText;
                                _originalTextMap[tb] = orig;
                            }

                            string target = TrText(orig);
                            if (tb.Text != target) tb.Text = target;
                            _lastAppliedTextMap[tb] = target;
                        }
                    }
                }
                else if (root is ComboBoxItem cbi && !_structurallyTranslatedElements.Contains(cbi) && !ReferenceEquals(ItemsControl.ItemsControlFromItemContainer(cbi), _cmbAppLanguage))
                {
                    if (cbi.Content is string cbiStr && !string.IsNullOrWhiteSpace(cbiStr))
                    {
                        if (!_originalTextMap.TryGetValue(cbi, out string orig) ||
                            (_lastAppliedTextMap.TryGetValue(cbi, out string lastApp) && cbiStr != lastApp))
                        {
                            orig = cbiStr;
                            _originalTextMap[cbi] = orig;
                        }
                        string target = TrText(orig);
                        if (!Equals(cbi.Content, target)) cbi.Content = target;
                        _lastAppliedTextMap[cbi] = target;
                    }
                }
                else if (root is TextBox searchTb && !string.IsNullOrWhiteSpace(searchTb.Text))
                {
                    if (!_originalTextMap.TryGetValue(searchTb, out string orig) ||
                        (_lastAppliedTextMap.TryGetValue(searchTb, out string lastApp) && searchTb.Text != lastApp))
                    {
                        orig = searchTb.Text;
                        _originalTextMap[searchTb] = orig;
                    }
                    string target = TrText(orig);
                    if (searchTb.Text != target && (orig.Contains("ara") || orig.Contains("Search") || orig.Contains("suchen") || orig.Contains("Rechercher") || orig.Contains("Buscar")))
                    {
                        searchTb.Text = target;
                        _lastAppliedTextMap[searchTb] = target;
                    }
                }
                else if (root is ContentControl cc && !_structurallyTranslatedElements.Contains(cc) && cc.Content is string strContent && cc is not ComboBoxItem && cc is not System.Windows.Controls.Primitives.ButtonBase && !string.IsNullOrWhiteSpace(strContent))
                {
                    if (!_originalTextMap.TryGetValue(cc, out string orig) ||
                        (_lastAppliedTextMap.TryGetValue(cc, out string lastApp) && strContent != lastApp))
                    {
                        orig = strContent;
                        _originalTextMap[cc] = orig;
                    }
                    string target = TrText(orig);
                    if (!Equals(cc.Content, target)) cc.Content = target;
                    _lastAppliedTextMap[cc] = target;
                }

                if (root is FrameworkElement fe && fe.ToolTip is string tipStr && !string.IsNullOrWhiteSpace(tipStr))
                {
                    if (!_originalToolTipMap.TryGetValue(fe, out string origTip) ||
                        (_lastAppliedToolTipMap.TryGetValue(fe, out string lastTip) && tipStr != lastTip))
                    {
                        origTip = tipStr;
                        _originalToolTipMap[fe] = origTip;
                    }
                    string targetTip = TrText(origTip);
                    if (!Equals(fe.ToolTip, targetTip)) fe.ToolTip = targetTip;
                    _lastAppliedToolTipMap[fe] = targetTip;
                }

                if (root is ComboBox combo && !ReferenceEquals(combo, _cmbAppLanguage))
                {
                    foreach (var item in combo.Items.OfType<DependencyObject>())
                    {
                        TranslateElementTree(item, visited);
                    }
                    RefreshComboBoxSelectionDisplay(combo);
                }

                if (root is Visual || root is System.Windows.Media.Media3D.Visual3D)
                {
                    int vCount = VisualTreeHelper.GetChildrenCount(root);
                    for (int i = 0; i < vCount; i++)
                    {
                        var child = VisualTreeHelper.GetChild(root, i);
                        if (child != null)
                            TranslateElementTree(child, visited);
                    }
                }

                if (root is Panel panel)
                {
                    foreach (UIElement child in panel.Children)
                    {
                        if (child != null)
                            TranslateElementTree(child, visited);
                    }
                }
                else if (root is Decorator dec && dec.Child != null)
                {
                    TranslateElementTree(dec.Child, visited);
                }
                else if (root is ContentControl contentCtrl && contentCtrl.Content is DependencyObject contentObj)
                {
                    TranslateElementTree(contentObj, visited);
                }
                else if (root is ItemsControl itemsCtrl)
                {
                    foreach (var item in itemsCtrl.Items)
                    {
                        if (item is DependencyObject itemDep)
                            TranslateElementTree(itemDep, visited);
                    }
                }
            }
            catch { }
        }
        // #endregion SARPNOTE_LOCALIZATION_ENGINE

        private static InkCanvas FindParentInkCanvas(DependencyObject src)
        {
            while (src != null)
            {
                if (src is InkCanvas ic) return ic;
                src = System.Windows.Media.VisualTreeHelper.GetParent(src);
            }
            return null;
        }

        private static Button FindParentButton(DependencyObject src)
        {
            while (src != null)
            {
                if (src is Button b) return b;
                src = System.Windows.Media.VisualTreeHelper.GetParent(src);
            }
            return null;
        }

        private void DualSelect_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Kullanıcı yan dock'ta başka bir araç butonuna (El, Kalem, Silgi vb.) bastıysa Seçim aracını anında bırak
            var clickedBtn = FindParentButton(e.OriginalSource as DependencyObject);
            if (clickedBtn != null && clickedBtn != BtnLasso && clickedBtn != BtnPickLassoMode && clickedBtn != BtnPickRectMode)
            {
                DependencyObject dockRoot = BtnLasso != null ? System.Windows.Media.VisualTreeHelper.GetParent(BtnLasso) : null;
                if (dockRoot != null && System.Windows.Media.VisualTreeHelper.GetParent(dockRoot) is Panel pPanel)
                    dockRoot = pPanel;

                if (dockRoot != null && IsDescendantOf(clickedBtn, dockRoot))
                {
                    _isSelectToolExclusiveActive = false;
                    RestoreOtherDockButtonsVisuals();
                    if (BtnLasso != null) BtnLasso.Tag = null;
                    if (SelectModePopup != null) SelectModePopup.IsOpen = false;

                    foreach (var ic in EnumerateRealizedInkCanvases(PagesStack))
                    {
                        ic.Select(new StrokeCollection());
                    }
                    return;
                }
            }

            if (!_isSelectToolExclusiveActive || _selectSubMode != "Rect") return;

            var ink = FindParentInkCanvas(e.OriginalSource as DependencyObject);
            if (ink == null || ink.EditingMode != InkCanvasEditingMode.Select) return;

            if (SelectModePopup != null && SelectModePopup.IsOpen)
                SelectModePopup.IsOpen = false;

            Point pt = e.GetPosition(ink);

            if (ink.GetSelectedStrokes().Count > 0)
            {
                Rect bounds = ink.GetSelectionBounds();
                if (!bounds.IsEmpty)
                {
                    bounds.Inflate(16, 16);
                    if (bounds.Contains(pt)) return;
                }
            }

            ink.Select(new StrokeCollection());
            _isRectSelecting = true;
            _rectSelectInk = ink;
            _rectSelectStart = pt;
            ink.EditingMode = InkCanvasEditingMode.None;

            var layer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(ink);
            if (layer != null)
            {
                _rectSelectAdorner = new RectSelectionAdorner(ink, pt);
                layer.Add(_rectSelectAdorner);
            }

            ink.CaptureMouse();
            e.Handled = true;
        }

        private static bool IsDescendantOf(DependencyObject node, DependencyObject ancestor)
        {
            while (node != null)
            {
                if (ReferenceEquals(node, ancestor)) return true;
                node = System.Windows.Media.VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        private void DualSelect_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isRectSelecting || _rectSelectInk == null) return;
            Point cur = e.GetPosition(_rectSelectInk);
            _rectSelectAdorner?.UpdatePoint(cur);
            e.Handled = true;
        }

        private void DualSelect_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isRectSelecting || _rectSelectInk == null) return;

            var ink = _rectSelectInk;
            Point endPt = e.GetPosition(ink);

            try
            {
                if (_rectSelectAdorner != null)
                {
                    var layer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(ink);
                    layer?.Remove(_rectSelectAdorner);
                    _rectSelectAdorner = null;
                }

                if (ink.IsMouseCaptured) ink.ReleaseMouseCapture();
                ink.EditingMode = InkCanvasEditingMode.Select;

                Rect box = new Rect(_rectSelectStart, endPt);
                if (box.Width > 3 || box.Height > 3)
                {
                    var matched = new StrokeCollection();
                    foreach (var st in ink.Strokes)
                    {
                        Rect sb = st.GetBounds();
                        if (!box.IntersectsWith(sb)) continue;
                        if (box.Contains(sb) || st.HitTest(box, 1) || st.StylusPoints.Any(sp => box.Contains(new Point(sp.X, sp.Y))))
                        {
                            matched.Add(st);
                        }
                    }
                    if (matched.Count > 0)
                    {
                        ink.Select(matched);
                    }
                }
            }
            finally
            {
                _isRectSelecting = false;
                _rectSelectInk = null;
                e.Handled = true;
            }
        }
        // #endregion DUAL_SELECTION_ENGINE
    }
    // #region PDF_CRISP_EXTENSION
    internal static class PdfCrispRenderExtensions
    {
        public static double CurrentZoomFactor = 1.30;

                public static System.Drawing.Image RenderCrisp(this PdfDocument doc, int page, int width, int height, float dpiX, float dpiY, PdfRenderFlags flags)
        {
            if (width <= 300 || height <= 350)
                return doc.Render(page, width, height, dpiX, dpiY, flags);

            var sz = doc.PageSizes[page];
            double rawW = Math.Max(1.0, sz.Width);
            double rawH = Math.Max(1.0, sz.Height);
            int targetW = Math.Clamp(width > 400 ? width : (int)Math.Round((rawH >= rawW ? 860.0 : 1180.0) * 2.1), 1200, 3200);
            int targetH = Math.Max(1, (int)Math.Round(targetW * (rawH / rawW)));
            float targetDpi = Math.Clamp(dpiX >= 96f ? dpiX : 144f, 110f, 220f);

            return doc.Render(page, targetW, targetH, targetDpi, targetDpi, PdfRenderFlags.Annotations);
        }

        public static System.Drawing.Image RenderCrisp(this PdfDocument doc, int page, int width, int height, float dpiX, float dpiY, bool forPrinting)
        {
            return RenderCrisp(doc, page, width, height, dpiX, dpiY, forPrinting ? PdfRenderFlags.ForPrinting : PdfRenderFlags.Annotations);
        }
    }
    // #endregion PDF_CRISP_EXTENSION
}






















        

