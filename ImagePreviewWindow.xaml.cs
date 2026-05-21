using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingImage = System.Drawing.Image;

namespace WordPictureAddIn
{
    public partial class ImagePreviewWindow : Window
    {
        private const double ZoomStepFactor = 1.06d;
        private const double AnnotationStrokeThickness = 3.0d;

        private DrawingBitmap _bitmap;
        private double _baseImageWidth;
        private double _baseImageHeight;
        private double _zoom = 1.0d;
        private double _initialScale = 1.0d;
        private double _initialTranslateX;
        private double _initialTranslateY;
        private bool _draggingImage;
        private bool _drawingAnnotation;
        private Point _dragStart;
        private double _dragOriginX;
        private double _dragOriginY;
        private Point _annotationStart;
        private Shape _previewShape;
        private Path _previewArrow;
        private AnnotationTool _currentTool = AnnotationTool.None;
        private readonly List<UIElement> _annotationHistory = new List<UIElement>();
        private Grid _activeTextEditorContainer;
        private TextBox _activeTextEditor;
        private bool _suspendAutoHide;
        private bool _isClosing;

        public ImagePreviewWindow()
        {
            InitializeComponent();
            Loaded += ImagePreviewWindow_Loaded;
            SizeChanged += ImagePreviewWindow_SizeChanged;
            Deactivated += ImagePreviewWindow_Deactivated;
            UpdateCursorVisual();
        }

        public void SetImage(DrawingImage image)
        {
            var old = _bitmap;
            _bitmap = new DrawingBitmap(image);
            if (old != null)
            {
                old.Dispose();
            }

            using (var stream = new System.IO.MemoryStream())
            {
                _bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;

                var source = new BitmapImage();
                source.BeginInit();
                source.CacheOption = BitmapCacheOption.OnLoad;
                source.StreamSource = stream;
                source.EndInit();
                source.Freeze();
                PreviewImage.Source = source;
            }

            WindowState = WindowState.Normal;
            Left = SystemParameters.VirtualScreenLeft;
            Top = SystemParameters.VirtualScreenTop;
            Width = SystemParameters.VirtualScreenWidth;
            Height = SystemParameters.VirtualScreenHeight;
            WindowState = WindowState.Maximized;
            ClearAnnotations();
            ResetImageLayout();
        }

        private void ImagePreviewWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ResetImageLayout();
        }

        private void ImagePreviewWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!_draggingImage)
            {
                ResetImageLayout();
            }
        }

        private void ImagePreviewWindow_Deactivated(object sender, EventArgs e)
        {
            if (_suspendAutoHide || _isClosing)
            {
                return;
            }

            ClearAnnotations();
            _isClosing = true;
            Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.Z)
            {
                CommitOrCancelTextEditing();
                UndoLastAnnotation();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                if (_currentTool != AnnotationTool.None)
                {
                    CancelTextEditing();
                    ClearToolSelection();
                }
                else
                {
                    if (_isClosing)
                    {
                        return;
                    }

                    ClearAnnotations();
                    _isClosing = true;
                    Close();
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var point = e.GetPosition(RootGrid);
            if (IsPointInsideActiveTextEditor(point))
            {
                return;
            }

            CommitOrCancelTextEditing();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing)
            {
                return;
            }

            CommitOrCancelTextEditing();
            ClearAnnotations();
            _isClosing = true;
            Close();
        }

        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();
            UndoLastAnnotation();
        }

        private void ResetZoomButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();
            RestoreInitialView();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();
            ClearToolSelection();

            if (_bitmap == null)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "PNG Image|*.png|JPEG Image|*.jpg;*.jpeg|BMP Image|*.bmp",
                AddExtension = true,
                DefaultExt = ".png",
                FileName = "word-preview.png"
            };

            _suspendAutoHide = true;
            try
            {
                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                using (var savedBitmap = BuildSavedBitmap())
                {
                    if (savedBitmap == null)
                    {
                        return;
                    }

                    var extension = System.IO.Path.GetExtension(dialog.FileName).ToLowerInvariant();
                    var imageFormat = System.Drawing.Imaging.ImageFormat.Png;
                    switch (extension)
                    {
                        case ".jpg":
                        case ".jpeg":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                            break;
                        case ".bmp":
                            imageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
                            break;
                    }

                    savedBitmap.Save(dialog.FileName, imageFormat);
                }
            }
            finally
            {
                _suspendAutoHide = false;
                Activate();
            }
        }

        private DrawingBitmap BuildSavedBitmap()
        {
            if (_bitmap == null || _baseImageWidth <= 0 || _baseImageHeight <= 0)
            {
                return null;
            }

            var output = new DrawingBitmap(_bitmap);
            var scaleX = output.Width / _baseImageWidth;
            var scaleY = output.Height / _baseImageHeight;

            using (var graphics = System.Drawing.Graphics.FromImage(output))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                foreach (var element in _annotationHistory)
                {
                    DrawAnnotationElement(graphics, element, scaleX, scaleY);
                }
            }

            return output;
        }

        private void DrawAnnotationElement(System.Drawing.Graphics graphics, UIElement element, double scaleX, double scaleY)
        {
            if (element is Line line)
            {
                using (var pen = CreateDrawingPen(scaleX))
                {
                    graphics.DrawLine(
                        pen,
                        (float)(line.X1 * scaleX),
                        (float)(line.Y1 * scaleY),
                        (float)(line.X2 * scaleX),
                        (float)(line.Y2 * scaleY));
                }

                return;
            }

            if (element is Rectangle rectangle)
            {
                using (var pen = CreateDrawingPen(scaleX))
                {
                    var left = Canvas.GetLeft(rectangle) * scaleX;
                    var top = Canvas.GetTop(rectangle) * scaleY;
                    var width = rectangle.Width * scaleX;
                    var height = rectangle.Height * scaleY;
                    graphics.DrawRectangle(pen, (float)left, (float)top, (float)width, (float)height);
                }

                return;
            }

            if (element is Ellipse ellipse)
            {
                using (var pen = CreateDrawingPen(scaleX))
                {
                    var left = Canvas.GetLeft(ellipse) * scaleX;
                    var top = Canvas.GetTop(ellipse) * scaleY;
                    var width = ellipse.Width * scaleX;
                    var height = ellipse.Height * scaleY;
                    graphics.DrawEllipse(pen, (float)left, (float)top, (float)width, (float)height);
                }

                return;
            }

            if (element is Path path)
            {
                DrawPathAnnotation(graphics, path, scaleX, scaleY);
                return;
            }

            if (element is Grid grid)
            {
                DrawTextAnnotation(graphics, grid, scaleX, scaleY);
            }
        }

        private void DrawPathAnnotation(System.Drawing.Graphics graphics, Path path, double scaleX, double scaleY)
        {
            var geometry = path.Data as GeometryGroup;
            if (geometry == null)
            {
                return;
            }

            using (var pen = CreateDrawingPen(scaleX))
            {
                foreach (var child in geometry.Children)
                {
                    if (child is LineGeometry line)
                    {
                        graphics.DrawLine(
                            pen,
                            (float)(line.StartPoint.X * scaleX),
                            (float)(line.StartPoint.Y * scaleY),
                            (float)(line.EndPoint.X * scaleX),
                            (float)(line.EndPoint.Y * scaleY));
                    }
                }
            }
        }

        private void DrawTextAnnotation(System.Drawing.Graphics graphics, Grid grid, double scaleX, double scaleY)
        {
            if (grid.Children.Count == 0 || !(grid.Children[0] is TextBlock textBlock))
            {
                return;
            }

            var left = Canvas.GetLeft(grid);
            var top = Canvas.GetTop(grid);
            var width = grid.Width;

            using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.OrangeRed))
            using (var font = new System.Drawing.Font("Microsoft YaHei UI", (float)(24 * scaleY), System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel))
            {
                var layoutRect = new System.Drawing.RectangleF(
                    (float)((left + 8) * scaleX),
                    (float)((top + 6) * scaleY),
                    (float)Math.Max(1, (width - 16) * scaleX),
                    10000f);

                var format = new System.Drawing.StringFormat
                {
                    Alignment = System.Drawing.StringAlignment.Near,
                    LineAlignment = System.Drawing.StringAlignment.Near
                };

                graphics.DrawString(textBlock.Text, font, brush, layoutRect, format);
            }
        }

        private static System.Drawing.Pen CreateDrawingPen(double scale)
        {
            var pen = new System.Drawing.Pen(System.Drawing.Color.OrangeRed, (float)Math.Max(1.0d, AnnotationStrokeThickness * scale))
            {
                StartCap = System.Drawing.Drawing2D.LineCap.Round,
                EndCap = System.Drawing.Drawing2D.LineCap.Round,
                LineJoin = System.Drawing.Drawing2D.LineJoin.Round
            };
            return pen;
        }

        private void ToolButton_Click(object sender, RoutedEventArgs e)
        {
            CommitOrCancelTextEditing();

            var clicked = sender as ToggleButton;
            if (clicked == null)
            {
                return;
            }

            var isChecked = clicked.IsChecked == true;
            SetToolButtonState(LineToolButton, clicked == LineToolButton && isChecked);
            SetToolButtonState(RectangleToolButton, clicked == RectangleToolButton && isChecked);
            SetToolButtonState(EllipseToolButton, clicked == EllipseToolButton && isChecked);
            SetToolButtonState(ArrowToolButton, clicked == ArrowToolButton && isChecked);
            SetToolButtonState(TextToolButton, clicked == TextToolButton && isChecked);

            if (!isChecked)
            {
                _currentTool = AnnotationTool.None;
                AnnotationCanvas.IsHitTestVisible = false;
                UpdateCursorVisual();
                return;
            }

            switch ((string)clicked.Tag)
            {
                case "Line":
                    _currentTool = AnnotationTool.Line;
                    break;
                case "Rectangle":
                    _currentTool = AnnotationTool.Rectangle;
                    break;
                case "Ellipse":
                    _currentTool = AnnotationTool.Ellipse;
                    break;
                case "Arrow":
                    _currentTool = AnnotationTool.Arrow;
                    break;
                case "Text":
                    _currentTool = AnnotationTool.Text;
                    break;
                default:
                    _currentTool = AnnotationTool.None;
                    break;
            }

            AnnotationCanvas.IsHitTestVisible = _currentTool != AnnotationTool.None;
            UpdateCursorVisual();
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_bitmap == null)
            {
                return;
            }

            var factor = e.Delta > 0 ? ZoomStepFactor : 1.0d / ZoomStepFactor;
            ChangeZoom(factor, e.GetPosition(ImageCanvas));
            e.Handled = true;
        }

        private void PreviewImage_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            Window_MouseWheel(sender, e);
        }

        private void PreviewImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool != AnnotationTool.None)
            {
                return;
            }

            _draggingImage = true;
            _dragStart = e.GetPosition(ImageCanvas);
            _dragOriginX = ContentTranslateTransform.X;
            _dragOriginY = ContentTranslateTransform.Y;
            ContentCanvas.CaptureMouse();
            e.Handled = true;
        }

        private void PreviewImage_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_draggingImage || _currentTool != AnnotationTool.None)
            {
                return;
            }

            var point = e.GetPosition(ImageCanvas);
            ContentTranslateTransform.X = _dragOriginX + (point.X - _dragStart.X);
            ContentTranslateTransform.Y = _dragOriginY + (point.Y - _dragStart.Y);
            ClampTranslation();
        }

        private void PreviewImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            _draggingImage = false;
            ContentCanvas.ReleaseMouseCapture();
            e.Handled = true;
        }

        private void AnnotationCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool == AnnotationTool.None)
            {
                return;
            }

            var point = ToContentCoordinates(e.GetPosition(ImageCanvas));

            if (_currentTool == AnnotationTool.Text)
            {
                AddTextAnnotation(point);
                e.Handled = true;
                return;
            }

            _drawingAnnotation = true;
            _annotationStart = point;
            AnnotationCanvas.CaptureMouse();
            CreatePreviewAnnotation(point);
            e.Handled = true;
        }

        private void AnnotationCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_drawingAnnotation)
            {
                return;
            }

            var point = ToContentCoordinates(e.GetPosition(ImageCanvas));
            UpdatePreviewAnnotation(_annotationStart, point);
        }

        private void AnnotationCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_drawingAnnotation)
            {
                return;
            }

            _drawingAnnotation = false;
            AnnotationCanvas.ReleaseMouseCapture();

            var point = ToContentCoordinates(e.GetPosition(ImageCanvas));
            CommitPreviewAnnotation(_annotationStart, point);
            RemovePreviewAnnotation();
            UpdateCursorVisual();
            e.Handled = true;
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            var point = e.GetPosition(RootGrid);

            if (IsPointInsideToolbar(point) || IsPointInsideCloseButton(point))
            {
                CursorOverlay.Visibility = Visibility.Collapsed;
                return;
            }

            CursorOverlay.Visibility = Visibility.Visible;
            UpdateCursorPosition(point);
        }

        private void ResetImageLayout()
        {
            if (_bitmap == null || ActualWidth <= 0 || ActualHeight <= 0)
            {
                return;
            }

            var usableWidth = Math.Max(100.0d, ActualWidth - 80.0d);
            var usableHeight = Math.Max(100.0d, ActualHeight - 80.0d);
            var scaleX = usableWidth / _bitmap.Width;
            var scaleY = usableHeight / _bitmap.Height;
            _zoom = Math.Min(1.0d, Math.Min(scaleX, scaleY));
            _zoom = Math.Max(0.1d, _zoom);

            _baseImageWidth = _bitmap.Width * _zoom;
            _baseImageHeight = _bitmap.Height * _zoom;
            PreviewImage.Width = _baseImageWidth;
            PreviewImage.Height = _baseImageHeight;
            AnnotationCanvas.Width = _baseImageWidth;
            AnnotationCanvas.Height = _baseImageHeight;
            ContentCanvas.Width = _baseImageWidth;
            ContentCanvas.Height = _baseImageHeight;
            ContentScaleTransform.ScaleX = 1.0d;
            ContentScaleTransform.ScaleY = 1.0d;
            ContentTranslateTransform.X = (ActualWidth - _baseImageWidth) / 2.0d;
            ContentTranslateTransform.Y = (ActualHeight - _baseImageHeight) / 2.0d;

            _initialScale = 1.0d;
            _initialTranslateX = ContentTranslateTransform.X;
            _initialTranslateY = ContentTranslateTransform.Y;
        }

        private void ChangeZoom(double factor, Point cursorPoint)
        {
            var oldScale = ContentScaleTransform.ScaleX;
            var newScale = Math.Max(0.1d, Math.Min(8.0d, oldScale * factor));

            var actualWidth = _baseImageWidth * oldScale;
            var actualHeight = _baseImageHeight * oldScale;
            if (actualWidth <= 0 || actualHeight <= 0)
            {
                return;
            }

            var relativeX = (cursorPoint.X - ContentTranslateTransform.X) / actualWidth;
            var relativeY = (cursorPoint.Y - ContentTranslateTransform.Y) / actualHeight;

            ContentScaleTransform.ScaleX = newScale;
            ContentScaleTransform.ScaleY = newScale;

            var newWidth = _baseImageWidth * newScale;
            var newHeight = _baseImageHeight * newScale;
            ContentTranslateTransform.X = cursorPoint.X - (newWidth * relativeX);
            ContentTranslateTransform.Y = cursorPoint.Y - (newHeight * relativeY);
            ClampTranslation();
        }

        private void ClampTranslation()
        {
            var imageWidth = _baseImageWidth * ContentScaleTransform.ScaleX;
            var imageHeight = _baseImageHeight * ContentScaleTransform.ScaleY;

            if (imageWidth <= ActualWidth)
            {
                ContentTranslateTransform.X = (ActualWidth - imageWidth) / 2.0d;
            }
            else
            {
                var minX = ActualWidth - imageWidth - 24.0d;
                var maxX = 24.0d;
                ContentTranslateTransform.X = Math.Max(minX, Math.Min(maxX, ContentTranslateTransform.X));
            }

            if (imageHeight <= ActualHeight)
            {
                ContentTranslateTransform.Y = (ActualHeight - imageHeight) / 2.0d;
            }
            else
            {
                var minY = ActualHeight - imageHeight - 24.0d;
                var maxY = 24.0d;
                ContentTranslateTransform.Y = Math.Max(minY, Math.Min(maxY, ContentTranslateTransform.Y));
            }
        }

        private void SetToolButtonState(ToggleButton button, bool isChecked)
        {
            button.IsChecked = isChecked;
        }

        private void ClearToolSelection()
        {
            _drawingAnnotation = false;
            _currentTool = AnnotationTool.None;
            AnnotationCanvas.IsHitTestVisible = false;
            AnnotationCanvas.ReleaseMouseCapture();
            RemovePreviewAnnotation();
            SetToolButtonState(LineToolButton, false);
            SetToolButtonState(RectangleToolButton, false);
            SetToolButtonState(EllipseToolButton, false);
            SetToolButtonState(ArrowToolButton, false);
            SetToolButtonState(TextToolButton, false);
            UpdateCursorVisual();
        }

        private void RestoreInitialView()
        {
            ContentScaleTransform.ScaleX = _initialScale;
            ContentScaleTransform.ScaleY = _initialScale;
            ContentTranslateTransform.X = _initialTranslateX;
            ContentTranslateTransform.Y = _initialTranslateY;
        }

        private Point ToContentCoordinates(Point canvasPoint)
        {
            var scaleX = ContentScaleTransform.ScaleX;
            var scaleY = ContentScaleTransform.ScaleY;

            if (scaleX == 0.0d || scaleY == 0.0d)
            {
                return new Point(0, 0);
            }

            return new Point(
                (canvasPoint.X - ContentTranslateTransform.X) / scaleX,
                (canvasPoint.Y - ContentTranslateTransform.Y) / scaleY);
        }

        private void CreatePreviewAnnotation(Point point)
        {
            RemovePreviewAnnotation();

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    _previewShape = new Line
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        X1 = point.X,
                        Y1 = point.Y,
                        X2 = point.X,
                        Y2 = point.Y
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Rectangle:
                    _previewShape = new Rectangle
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Ellipse:
                    _previewShape = new Ellipse
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewShape);
                    break;
                case AnnotationTool.Arrow:
                    _previewArrow = new Path
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(_previewArrow);
                    break;
            }
        }

        private void UpdatePreviewAnnotation(Point start, Point end)
        {
            end = ApplyShiftConstraint(start, end, _currentTool);

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    var line = _previewShape as Line;
                    if (line != null)
                    {
                        line.X1 = start.X;
                        line.Y1 = start.Y;
                        line.X2 = end.X;
                        line.Y2 = end.Y;
                    }
                    break;
                case AnnotationTool.Rectangle:
                    UpdateSizedShape(_previewShape, start, end);
                    break;
                case AnnotationTool.Ellipse:
                    UpdateSizedShape(_previewShape, start, end);
                    break;
                case AnnotationTool.Arrow:
                    if (_previewArrow != null)
                    {
                        _previewArrow.Data = BuildArrowGeometry(start, end);
                    }
                    break;
            }
        }

        private void CommitPreviewAnnotation(Point start, Point end)
        {
            end = ApplyShiftConstraint(start, end, _currentTool);

            if (Math.Abs(end.X - start.X) < 2 && Math.Abs(end.Y - start.Y) < 2)
            {
                return;
            }

            switch (_currentTool)
            {
                case AnnotationTool.Line:
                    var line = new Line
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        X1 = start.X,
                        Y1 = start.Y,
                        X2 = end.X,
                        Y2 = end.Y
                    };
                    AnnotationCanvas.Children.Add(line);
                    _annotationHistory.Add(line);
                    break;
                case AnnotationTool.Rectangle:
                    var rect = new Rectangle
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(rect);
                    UpdateSizedShape(rect, start, end);
                    _annotationHistory.Add(rect);
                    break;
                case AnnotationTool.Ellipse:
                    var ellipse = new Ellipse
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent
                    };
                    AnnotationCanvas.Children.Add(ellipse);
                    UpdateSizedShape(ellipse, start, end);
                    _annotationHistory.Add(ellipse);
                    break;
                case AnnotationTool.Arrow:
                    var arrow = new Path
                    {
                        Stroke = Brushes.OrangeRed,
                        StrokeThickness = AnnotationStrokeThickness,
                        Fill = Brushes.Transparent,
                        Data = BuildArrowGeometry(start, end)
                    };
                    AnnotationCanvas.Children.Add(arrow);
                    _annotationHistory.Add(arrow);
                    break;
            }
        }

        private void UpdateSizedShape(Shape shape, Point start, Point end)
        {
            if (shape == null)
            {
                return;
            }

            var left = Math.Min(start.X, end.X);
            var top = Math.Min(start.Y, end.Y);
            var width = Math.Abs(end.X - start.X);
            var height = Math.Abs(end.Y - start.Y);

            Canvas.SetLeft(shape, left);
            Canvas.SetTop(shape, top);
            shape.Width = width;
            shape.Height = height;
        }

        private Geometry BuildArrowGeometry(Point start, Point end)
        {
            var group = new GeometryGroup();
            group.Children.Add(new LineGeometry(start, end));

            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 0.1d)
            {
                return group;
            }

            var ux = dx / length;
            var uy = dy / length;
            var arrowLength = 14.0d;
            var arrowWidth = 6.0d;

            var left = new Point(
                end.X - arrowLength * ux + arrowWidth * uy,
                end.Y - arrowLength * uy - arrowWidth * ux);
            var right = new Point(
                end.X - arrowLength * ux - arrowWidth * uy,
                end.Y - arrowLength * uy + arrowWidth * ux);

            group.Children.Add(new LineGeometry(end, left));
            group.Children.Add(new LineGeometry(end, right));
            return group;
        }

        private Point ApplyShiftConstraint(Point start, Point end, AnnotationTool tool)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) != ModifierKeys.Shift)
            {
                return end;
            }

            switch (tool)
            {
                case AnnotationTool.Line:
                case AnnotationTool.Arrow:
                    return SnapLineToEightDirections(start, end);
                case AnnotationTool.Rectangle:
                case AnnotationTool.Ellipse:
                    return SnapToSquare(start, end);
                default:
                    return end;
            }
        }

        private static Point SnapLineToEightDirections(Point start, Point end)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 0.001d)
            {
                return end;
            }

            var angle = Math.Atan2(dy, dx);
            var step = Math.PI / 4.0d;
            var snappedAngle = Math.Round(angle / step) * step;

            return new Point(
                start.X + Math.Cos(snappedAngle) * length,
                start.Y + Math.Sin(snappedAngle) * length);
        }

        private static Point SnapToSquare(Point start, Point end)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var size = Math.Max(Math.Abs(dx), Math.Abs(dy));

            return new Point(
                start.X + Math.Sign(dx == 0 ? 1 : dx) * size,
                start.Y + Math.Sign(dy == 0 ? 1 : dy) * size);
        }

        private void AddTextAnnotation(Point point)
        {
            CommitOrCancelTextEditing();

            var container = new Grid
            {
                Width = 220,
                Height = 56
            };

            var rectangle = new Rectangle
            {
                Stroke = Brushes.Gray,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(28, 255, 255, 255)),
                StrokeDashArray = new DoubleCollection { 4, 3 },
                RadiusX = 4,
                RadiusY = 4
            };

            var textBox = new TextBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Foreground = Brushes.OrangeRed,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(8, 6, 8, 6),
                CaretBrush = Brushes.OrangeRed
            };

            container.Children.Add(rectangle);
            container.Children.Add(textBox);
            AnnotationCanvas.Children.Add(container);
            Canvas.SetLeft(container, point.X);
            Canvas.SetTop(container, point.Y);

            _activeTextEditorContainer = container;
            _activeTextEditor = textBox;

            textBox.Loaded += delegate
            {
                textBox.Focus();
                textBox.Select(textBox.Text.Length, 0);
            };
            textBox.Focus();
            textBox.Select(textBox.Text.Length, 0);
        }

        private void CommitOrCancelTextEditing()
        {
            if (_activeTextEditorContainer == null || _activeTextEditor == null)
            {
                return;
            }

            var text = _activeTextEditor.Text;

            if (!string.IsNullOrWhiteSpace(text))
            {
                var textBlock = new TextBlock
                {
                    Text = text,
                    Foreground = Brushes.OrangeRed,
                    FontFamily = new FontFamily("Microsoft YaHei UI"),
                    FontSize = 24,
                    FontWeight = FontWeights.Bold,
                    TextWrapping = TextWrapping.Wrap,
                    Width = _activeTextEditorContainer.Width,
                    Margin = new Thickness(8, 6, 8, 6),
                    LineHeight = 28
                };

                _activeTextEditorContainer.Children.Clear();
                _activeTextEditorContainer.Children.Add(textBlock);
                _annotationHistory.Add(_activeTextEditorContainer);
            }
            else
            {
                AnnotationCanvas.Children.Remove(_activeTextEditorContainer);
            }

            _activeTextEditor = null;
            _activeTextEditorContainer = null;
        }

        private void CancelTextEditing()
        {
            if (_activeTextEditorContainer == null)
            {
                return;
            }

            AnnotationCanvas.Children.Remove(_activeTextEditorContainer);
            _activeTextEditor = null;
            _activeTextEditorContainer = null;
        }

        private void RemovePreviewAnnotation()
        {
            if (_previewShape != null)
            {
                AnnotationCanvas.Children.Remove(_previewShape);
                _previewShape = null;
            }

            if (_previewArrow != null)
            {
                AnnotationCanvas.Children.Remove(_previewArrow);
                _previewArrow = null;
            }
        }

        private void ClearAnnotations()
        {
            CancelTextEditing();
            AnnotationCanvas.Children.Clear();
            _annotationHistory.Clear();
            _previewShape = null;
            _previewArrow = null;
            _drawingAnnotation = false;
            ClearToolSelection();
        }

        private void UndoLastAnnotation()
        {
            if (_annotationHistory.Count == 0)
            {
                return;
            }

            var lastIndex = _annotationHistory.Count - 1;
            var element = _annotationHistory[lastIndex];
            _annotationHistory.RemoveAt(lastIndex);
            AnnotationCanvas.Children.Remove(element);
        }

        private bool IsPointInsideActiveTextEditor(Point rootPoint)
        {
            if (_activeTextEditorContainer == null)
            {
                return false;
            }

            var topLeft = _activeTextEditorContainer.TranslatePoint(new Point(0, 0), RootGrid);
            var rect = new Rect(topLeft.X, topLeft.Y, _activeTextEditorContainer.ActualWidth, _activeTextEditorContainer.ActualHeight);
            return rect.Contains(rootPoint);
        }

        private bool IsPointInsideCloseButton(Point point)
        {
            var topLeft = CloseButton.TranslatePoint(new Point(0, 0), RootGrid);
            var rect = new Rect(topLeft.X, topLeft.Y, CloseButton.ActualWidth, CloseButton.ActualHeight);
            return rect.Contains(point);
        }

        private bool IsPointInsideToolbar(Point point)
        {
            var toolbarTopLeft = ToolbarPanel.TranslatePoint(new Point(0, 0), RootGrid);
            var rect = new Rect(toolbarTopLeft.X, toolbarTopLeft.Y, ToolbarPanel.ActualWidth, ToolbarPanel.ActualHeight);
            return rect.Contains(point);
        }

        private void UpdateCursorVisual()
        {
            var isDrawMode = _currentTool != AnnotationTool.None;
            HandCursorVisual.Visibility = isDrawMode ? Visibility.Collapsed : Visibility.Visible;
            PenCursorVisual.Visibility = isDrawMode ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateCursorPosition(Point point)
        {
            var activeCursor = _currentTool != AnnotationTool.None ? (FrameworkElement)PenCursorVisual : HandCursorVisual;
            Canvas.SetLeft(activeCursor, point.X + 1);
            Canvas.SetTop(activeCursor, point.Y + 1);
        }

        private enum AnnotationTool
        {
            None,
            Line,
            Rectangle,
            Ellipse,
            Arrow,
            Text
        }
    }
}
