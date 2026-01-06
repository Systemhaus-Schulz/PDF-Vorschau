using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Documents;
using System.Windows.Threading;
using Microsoft.Win32;
using Microsoft.VisualBasic; // für InputBox
using PdfiumViewer;
using WpfImage = System.Windows.Controls.Image;
using DrawingImage = System.Drawing.Image;
using DrawingBitmap = System.Drawing.Bitmap;

namespace PDF_Vorschau
{
    public partial class MainWindow : Window
    {
        // Hilfsklasse für ListBox-Einträge: zeigt nur den Dateinamen, speichert aber den vollständigen Pfad
        private sealed class PdfItem
        {
            public string FullPath { get; }
            public string Name { get; }

            public PdfItem(string fullPath)
            {
                FullPath = fullPath;
                Name = Path.GetFileName(fullPath);
            }

            public override string ToString() => Name;
        }

        private double _zoom = 1.0;
        private FileSystemWatcher? _watcher;
        private string? _currentFolder;
        private string? _currentPdfPath;

        public MainWindow()
        {
            InitializeComponent();

            // kleines Startfenster anzeigen
            ShowStartupWindow();

            PreviewImage.MouseDown += (s, e) => PreviewImage.Focus();
        }

        // ----------------------- STARTFENSTER -----------------------

        private void ShowStartupWindow()
        {
            var splash = CreateStartupWindow();
            splash.Show();

            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2.5)
            };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                if (splash.IsVisible)
                    splash.Close();
            };
            timer.Start();
        }

        private Window CreateStartupWindow()
        {
            var splash = new Window
            {
                Width = 420,
                Height = 220,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.None,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(Color.FromRgb(24, 24, 28)),
                Topmost = true,
                ShowInTaskbar = false
            };

            var outerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(32, 32, 38)),
                CornerRadius = new CornerRadius(12),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0, 180, 220)),
                BorderThickness = new Thickness(1.5),
                Padding = new Thickness(16)
            };

            var stack = new StackPanel
            {
                Orientation = Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            var title = new TextBlock
            {
                Text = "Systemhaus Schulz",
                Foreground = Brushes.White,
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 8),
                TextAlignment = TextAlignment.Center
            };

            var subtitle = new TextBlock
            {
                Text = "PDF-Vorschau für IT-Dokumente",
                Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 190)),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 16),
                TextAlignment = TextAlignment.Center
            };

            var email = new TextBlock
            {
                Text = "info@systemhaus-schulz.de",
                Foreground = new SolidColorBrush(Color.FromRgb(200, 200, 210)),
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 4),
                TextAlignment = TextAlignment.Center
            };

            var urlText = new TextBlock
            {
                Foreground = new SolidColorBrush(Color.FromRgb(0, 200, 255)),
                FontSize = 12,
                TextAlignment = TextAlignment.Center
            };

            var hyperlink = new Hyperlink
            {
                NavigateUri = new Uri("https://www.systemhaus-schulz.de")
            };
            hyperlink.Inlines.Add("https://www.systemhaus-schulz.de");
            hyperlink.RequestNavigate += (s, e) =>
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = e.Uri.AbsoluteUri,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    // ignorieren
                }
            };
            urlText.Inlines.Add(hyperlink);

            var footer = new TextBlock
            {
                Text = "© Systemhaus Schulz",
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 130)),
                FontSize = 11,
                Margin = new Thickness(0, 16, 0, 0),
                TextAlignment = TextAlignment.Center
            };

            // Optional: Logo einbinden, wenn als Resource vorhanden
            // var logo = new Image
            // {
            //     Width = 96,
            //     Height = 96,
            //     Margin = new Thickness(0, 0, 0, 12),
            //     Source = new BitmapImage(new Uri("pack://application:,,,/Resources/systemhaus-logo.png"))
            // };
            // stack.Children.Add(logo);

            stack.Children.Add(title);
            stack.Children.Add(subtitle);
            stack.Children.Add(email);
            stack.Children.Add(urlText);
            stack.Children.Add(footer);

            outerBorder.Child = stack;
            splash.Content = outerBorder;

            return splash;
        }

        // ----------------------- ORDNER WÄHLEN -----------------------

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "PDF-Dateien (*.pdf)|*.pdf",
                Multiselect = false,
                Title = "Eine PDF im gewünschten Ordner auswählen"
            };

            if (dlg.ShowDialog() == true)
            {
                string? folder = Path.GetDirectoryName(dlg.FileName);
                if (!string.IsNullOrEmpty(folder))
                {
                    _currentFolder = folder;
                    FolderPathTextBox.Text = folder;
                    SetupWatcher(folder);
                    LoadPdfFiles();
                }
            }
        }

        private void SetupWatcher(string folder)
        {
            try
            {
                _watcher?.Dispose();

                _watcher = new FileSystemWatcher(folder, "*.pdf")
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
                };

                _watcher.Created += (_, __) => Dispatcher.Invoke(LoadPdfFiles);
                _watcher.Deleted += (_, __) => Dispatcher.Invoke(LoadPdfFiles);
                _watcher.Renamed += (_, __) => Dispatcher.Invoke(LoadPdfFiles);
                _watcher.Changed += (_, __) => Dispatcher.Invoke(LoadPdfFiles);

                _watcher.EnableRaisingEvents = true;
            }
            catch
            {
                _watcher = null;
            }
        }

        private void LoadPdfFiles()
        {
            PdfListBox.Items.Clear();

            string path = FolderPathTextBox.Text;
            if (!Directory.Exists(path))
                return;

            if (_currentFolder == null || !string.Equals(_currentFolder, path, StringComparison.OrdinalIgnoreCase))
            {
                _currentFolder = path;
                SetupWatcher(path);
            }

            foreach (var file in Directory.GetFiles(path, "*.pdf"))
            {
                PdfListBox.Items.Add(new PdfItem(file));
            }

            if (!string.IsNullOrEmpty(_currentPdfPath))
            {
                var match = PdfListBox.Items.Cast<PdfItem>()
                    .FirstOrDefault(f => string.Equals(f.FullPath, _currentPdfPath, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                    PdfListBox.SelectedItem = match;
            }
        }

        // ----------------------- LISTBOX AUSWAHL -----------------------

        private void PdfListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PdfListBox.SelectedItem is PdfItem item)
            {
                _currentPdfPath = item.FullPath;
                ShowPreview(item.FullPath);
                LoadThumbnails(item.FullPath);
            }
        }

        private void PdfListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PdfListBox.SelectedItem is PdfItem item)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = item.FullPath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF konnte nicht geöffnet werden:\n{ex.Message}",
                        "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // F2 für Umbenennen
        private void PdfListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F2)
            {
                RenameSelectedPdf();
                e.Handled = true;
            }
        }

        // ----------------------- UMBENENNEN -----------------------

        private void RenameSelectedPdf()
        {
            if (PdfListBox.SelectedItem is not PdfItem selected)
                return;

            string oldPath = selected.FullPath;

            if (!File.Exists(oldPath))
            {
                MessageBox.Show("Die ausgewählte Datei existiert nicht mehr.", "Hinweis",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                LoadPdfFiles();
                return;
            }

            string? folder = Path.GetDirectoryName(oldPath);
            string oldName = Path.GetFileNameWithoutExtension(oldPath);

            if (string.IsNullOrEmpty(folder))
                return;

            string newName = Interaction.InputBox(
                "Neuen Dateinamen (ohne .pdf) eingeben:",
                "PDF umbenennen",
                oldName
            );

            if (string.IsNullOrWhiteSpace(newName) || newName == oldName)
                return;

            if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show("Der Dateiname enthält ungültige Zeichen.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string newPath = Path.Combine(folder, newName + ".pdf");

            if (File.Exists(newPath))
            {
                MessageBox.Show("Eine Datei mit diesem Namen existiert bereits.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Vorschau lösen, damit kein Handle mehr offen ist
                PreviewImage.Source = null;
                ThumbPanel.Children.Clear();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                var attrs = File.GetAttributes(oldPath);
                if ((attrs & FileAttributes.ReadOnly) != 0)
                {
                    File.SetAttributes(oldPath, attrs & ~FileAttributes.ReadOnly);
                }

                if (_watcher != null)
                    _watcher.EnableRaisingEvents = false;

                File.Move(oldPath, newPath);

                _currentPdfPath = newPath;

                LoadPdfFiles();

                var match = PdfListBox.Items.Cast<PdfItem>()
                    .FirstOrDefault(f => string.Equals(f.FullPath, newPath, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                    PdfListBox.SelectedItem = match;

                ShowPreview(newPath);
                LoadThumbnails(newPath);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show(
                    "Die Datei konnte nicht umbenannt werden (Zugriff verweigert).\n\n" +
                    $"Pfad: {oldPath}\n\n" +
                    "Mögliche Ursachen:\n" +
                    "• Die PDF ist noch in einem anderen Programm geöffnet (Adobe Reader, Browser, Explorer-Vorschau).\n" +
                    "• Der Ordner erfordert erhöhte Rechte (z. B. Program Files).\n" +
                    "• Es bestehen keine Schreibrechte für diesen Benutzer.\n\n" +
                    $"Technische Details:\n{ex.Message}",
                    "Zugriff verweigert",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Umbenennen der Datei:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (_watcher != null)
                    _watcher.EnableRaisingEvents = true;
            }
        }

        // ----------------------- PDF PREVIEW -----------------------

        private void ShowPreview(string path)
        {
            try
            {
                using var document = PdfDocument.Load(path);
                using var img = document.Render(0, 1200, 1200, true);
                PreviewImage.Source = ConvertToBitmapSource(img);
                _zoom = 1.0;
                PreviewImage.LayoutTransform = new ScaleTransform(_zoom, _zoom);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Rendern der Vorschau:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                PreviewImage.Source = null;
            }
        }

        private BitmapSource ConvertToBitmapSource(DrawingImage img)
        {
            using var bmp = new DrawingBitmap(img);
            var h = bmp.GetHbitmap();
            var bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                h, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            return bs;
        }

        // ----------------------- THUMBNAILS -----------------------

        private void LoadThumbnails(string path)
        {
            ThumbPanel.Children.Clear();

            try
            {
                using var document = PdfDocument.Load(path);
                int pages = document.PageCount;

                for (int i = 0; i < pages; i++)
                {
                    using var img = document.Render(i, 200, 200, true);
                    var bmpSource = ConvertToBitmapSource(img);

                    var iv = new WpfImage
                    {
                        Source = bmpSource,
                        Width = 140,
                        Height = 180,
                        Margin = new Thickness(4),
                        Cursor = Cursors.Hand
                    };

                    int pageIndex = i;
                    string localPath = path;

                    iv.MouseDown += (s, e) =>
                    {
                        try
                        {
                            using var doc2 = PdfDocument.Load(localPath);
                            using var full = doc2.Render(pageIndex, 1200, 1200, true);
                            PreviewImage.Source = ConvertToBitmapSource(full);
                            _zoom = 1.0;
                            PreviewImage.LayoutTransform = new ScaleTransform(_zoom, _zoom);
                        }
                        catch (Exception ex2)
                        {
                            MessageBox.Show($"Fehler beim Rendern der Seite:\n{ex2.Message}",
                                "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    };

                    ThumbPanel.Children.Add(iv);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Laden der Thumbnails:\n{ex.Message}",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ----------------------- ZOOM -----------------------

        private void PreviewImage_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                return;

            if (e.Delta > 0)
                _zoom += 0.1;
            else
                _zoom -= 0.1;

            if (_zoom < 0.2) _zoom = 0.2;
            if (_zoom > 5.0) _zoom = 5.0;

            PreviewImage.LayoutTransform = new ScaleTransform(_zoom, _zoom);
            e.Handled = true;
        }

        // STRG + / - auf der Tastatur
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                return;

            if (e.Key == Key.Add || e.Key == Key.OemPlus)
                _zoom += 0.1;
            else if (e.Key == Key.Subtract || e.Key == Key.OemMinus)
                _zoom -= 0.1;
            else
                return;

            if (_zoom < 0.2) _zoom = 0.2;
            if (_zoom > 5.0) _zoom = 5.0;

            PreviewImage.LayoutTransform = new ScaleTransform(_zoom, _zoom);
            e.Handled = true;
        }

        // ----------------------- MEHRSEITENMODUS -----------------------

        private void ToggleThumbs_Checked(object sender, RoutedEventArgs e)
        {
            if (ThumbPanel == null || ThumbBarContainer == null || ThumbRow == null)
                return;

            bool show = ToggleThumbs.IsChecked == true;

            ThumbPanel.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            ThumbBarContainer.Visibility = show ? Visibility.Visible : Visibility.Collapsed;

            ThumbRow.Height = show
                ? new GridLength(180)
                : new GridLength(0);
        }
    }
}
