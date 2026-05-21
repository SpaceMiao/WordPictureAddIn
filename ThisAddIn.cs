using System;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace WordPictureAddIn
{
    public partial class ThisAddIn
    {
        private ImagePreviewWindow _previewWindow;
        private string _cachedSelectionKey;
        private Image _cachedSelectionImage;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            EnsureWpfApplication();
            EnsurePreviewWindow();
            Application.WindowSelectionChange += Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WindowSelectionChange -= Application_WindowSelectionChange;
            Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;

            if (_previewWindow != null)
            {
                _previewWindow.Close();
                _previewWindow = null;
            }

            ClearCachedSelectionImage();
        }

        private void Application_WindowSelectionChange(Selection selection)
        {
            string selectionKey;
            if (!TryGetSelectionKey(selection, out selectionKey))
            {
                ClearCachedSelectionImage();
                return;
            }

            if (selectionKey == _cachedSelectionKey && _cachedSelectionImage != null)
            {
                return;
            }

            var image = CaptureSelectionImage(selection);
            if (image == null)
            {
                ClearCachedSelectionImage();
                return;
            }

            ReplaceCachedSelectionImage(selectionKey, image);
        }

        private void Application_WindowBeforeDoubleClick(Selection selection, ref bool cancel)
        {
            try
            {
                string selectionKey;
                if (!TryGetSelectionKey(selection, out selectionKey))
                {
                    return;
                }

                var image = TryGetCachedSelectionImage(selectionKey) ?? CaptureSelectionImage(selection);
                if (image == null)
                {
                    return;
                }

                ShowPreview(image);
            }
            catch
            {
            }
        }

        private static bool HasSelectedImage(Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            return selection.InlineShapes.Count > 0 ||
                   (selection.ShapeRange != null && selection.ShapeRange.Count > 0);
        }

        private static bool TryGetSelectionKey(Selection selection, out string selectionKey)
        {
            selectionKey = null;

            if (!HasSelectedImage(selection))
            {
                return false;
            }

            if (selection.InlineShapes.Count > 0)
            {
                var inlineShape = selection.InlineShapes[1];
                selectionKey = string.Format(
                    "inline:{0}:{1}:{2}:{3}",
                    selection.Range.Start,
                    selection.Range.End,
                    inlineShape.Width,
                    inlineShape.Height);
                return true;
            }

            if (selection.ShapeRange != null && selection.ShapeRange.Count > 0)
            {
                var shape = selection.ShapeRange[1];
                selectionKey = string.Format(
                    "shape:{0}:{1}:{2}:{3}",
                    shape.Left,
                    shape.Top,
                    shape.Width,
                    shape.Height);
                return true;
            }

            return false;
        }

        private static Image CaptureSelectionImage(Selection selection)
        {
            return TryCaptureSelectionImageViaEnhMetaFileBits(selection);
        }

        private static Image TryCaptureSelectionImageViaEnhMetaFileBits(Selection selection)
        {
            try
            {
                if (selection == null)
                {
                    return null;
                }

                var bits = selection.EnhMetaFileBits as byte[];
                if (bits == null || bits.Length == 0)
                {
                    return null;
                }

                using (var stream = new MemoryStream(bits))
                using (var bitmap = new Bitmap(stream))
                {
                    if (selection.Type == WdSelectionType.wdSelectionInlineShape &&
                        selection.InlineShapes.Count > 0)
                    {
                        var inlineShape = selection.InlineShapes[1];
                        var cropped = CropInlineShapeBitmap(bitmap, inlineShape);
                        if (cropped != null)
                        {
                            return cropped;
                        }
                    }

                    return new Bitmap(bitmap);
                }
            }
            catch
            {
                return null;
            }
        }

        private static Image CropInlineShapeBitmap(Bitmap sourceBitmap, InlineShape inlineShape)
        {
            try
            {
                if (sourceBitmap == null || inlineShape == null)
                {
                    return null;
                }

                if (inlineShape.Height <= 0 || inlineShape.Width <= 0)
                {
                    return new Bitmap(sourceBitmap);
                }

                var targetWidth = (int)Math.Round(sourceBitmap.Height * (inlineShape.Width / inlineShape.Height));
                targetWidth = Math.Max(1, Math.Min(targetWidth, sourceBitmap.Width));

                var cropRectangle = new System.Drawing.Rectangle(0, 0, targetWidth, sourceBitmap.Height);
                var croppedBitmap = new Bitmap(cropRectangle.Width, cropRectangle.Height);

                using (var graphics = Graphics.FromImage(croppedBitmap))
                {
                    graphics.DrawImage(
                        sourceBitmap,
                        new System.Drawing.Rectangle(0, 0, croppedBitmap.Width, croppedBitmap.Height),
                        cropRectangle,
                        System.Drawing.GraphicsUnit.Pixel);
                }

                return croppedBitmap;
            }
            catch
            {
                return new Bitmap(sourceBitmap);
            }
        }

        private void ShowPreview(Image image)
        {
            EnsurePreviewWindow();

            _previewWindow.SetImage(image);

            if (!_previewWindow.IsVisible)
            {
                _previewWindow.Show();
            }
            else
            {
                _previewWindow.Activate();
            }

            _previewWindow.Activate();
        }

        private Image TryGetCachedSelectionImage(string selectionKey)
        {
            if (_cachedSelectionImage == null || selectionKey != _cachedSelectionKey)
            {
                return null;
            }

            return new Bitmap(_cachedSelectionImage);
        }

        private void ReplaceCachedSelectionImage(string selectionKey, Image image)
        {
            ClearCachedSelectionImage();
            _cachedSelectionKey = selectionKey;
            _cachedSelectionImage = image;
        }

        private void ClearCachedSelectionImage()
        {
            if (_cachedSelectionImage != null)
            {
                _cachedSelectionImage.Dispose();
                _cachedSelectionImage = null;
            }

            _cachedSelectionKey = null;
        }

        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };
            }
        }

        private void EnsurePreviewWindow()
        {
            if (_previewWindow == null)
            {
                _previewWindow = new ImagePreviewWindow();
                _previewWindow.Closed += delegate
                {
                    _previewWindow = null;
                };
            }
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
