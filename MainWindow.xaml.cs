using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Aspose.Words;
using Aspose.Pdf;
using MediaToolkit;
using MediaToolkit.Model;
using System.Windows.Media.Animation;
using System.Windows.Media;
using System.Data.SQLite;
using System.Collections.Generic;
using Spire.Doc;
using Spire.Doc.Documents;


namespace WpfApp2
{
    public partial class MainWindow : Window
    {
        private string _selectedImagePath;
        private string _selectedTextFilePath;
        private string _selectedMediaFilePath;
        private string _lastConvertedFilePath;

        public MainWindow()
        {
            InitializeComponent();
            DatabaseHelper.InitializeDatabase();
            MigrateOldHistory();
            LoadConversionHistory();
        }

        // Image Conversion Methods
        private void SelectImageButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif",
                Title = "Выберите изображение"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedImagePath = openFileDialog.FileName;
                SelectedImagePath.Text = _selectedImagePath;
                AnimateTextBlock(SelectedImagePath);
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedImagePath))
            {
                StatusTextBlock.Text = "Пожалуйста, выберите изображение.";
                return;
            }

            string selectedFormat = ((ComboBoxItem)FormatComboBox.SelectedItem)?.Tag.ToString();
            if (string.IsNullOrEmpty(selectedFormat))
            {
                StatusTextBlock.Text = "Пожалуйста, выберите формат.";
                return;
            }

            try
            {
                string newFilePath = Path.ChangeExtension(_selectedImagePath, selectedFormat);
                BitmapImage bitmap = new BitmapImage(new Uri(_selectedImagePath));

                using (FileStream stream = new FileStream(newFilePath, FileMode.Create))
                {
                    BitmapEncoder encoder = GetEncoder(selectedFormat);
                    if (encoder != null)
                    {
                        encoder.Frames.Add(BitmapFrame.Create(bitmap));
                        encoder.Save(stream);
                    }
                }

                StatusTextBlock.Text = $"Изображение успешно конвертировано в {newFilePath}.";
                AddToConversionHistory(newFilePath, selectedFormat);
                _lastConvertedFilePath = newFilePath;
                LoadConversionHistory();

                OpenFolderButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"Ошибка: {ex.Message}";
            }
        }

        // Text File Conversion Methods
        private void SelectTextFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt;*.pdf;*.docx;*.doc",
                Title = "Выберите текстовый файл"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedTextFilePath = openFileDialog.FileName;
                SelectedTextFilePath.Text = _selectedTextFilePath;
                AnimateTextBlock(SelectedTextFilePath);
            }
        }

        private void ConvertTextButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedTextFilePath))
            {
                TextStatusTextBlock.Text = "Пожалуйста, выберите текстовый файл.";
                return;
            }

            string selectedFormat = ((ComboBoxItem)TextFormatComboBox.SelectedItem)?.Tag.ToString();
            if (string.IsNullOrEmpty(selectedFormat))
            {
                TextStatusTextBlock.Text = "Пожалуйста, выберите формат.";
                return;
            }

            try
            {
                ConvertTextFile(selectedFormat);
                OpenFolderButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                TextStatusTextBlock.Text = $"Ошибка: {ex.Message}";
            }
        }

        private void ConvertTextFile(string selectedFormat)
        {
            string newFilePath = Path.ChangeExtension(_selectedTextFilePath, selectedFormat);

            switch (selectedFormat)
            {
                case "pdf":
                    var docPdf = new Aspose.Words.Document(_selectedTextFilePath);
                    docPdf.Save(newFilePath, Aspose.Words.SaveFormat.Pdf);
                    break;

                case "docx":
                    var docx = new Aspose.Words.Document(_selectedTextFilePath);
                    docx.Save(newFilePath, Aspose.Words.SaveFormat.Docx);
                    break;

                case "txt":
                    var txtDoc = new Aspose.Words.Document(_selectedTextFilePath);
                    txtDoc.Save(newFilePath, Aspose.Words.SaveFormat.Text);
                    break;

                case "doc":
                    var docDoc = new Aspose.Words.Document(_selectedTextFilePath);
                    docDoc.Save(newFilePath, Aspose.Words.SaveFormat.Doc);
                    break;

                default:
                    throw new InvalidOperationException("Неподдерживаемый формат.");
            }

            TextStatusTextBlock.Text = $"Текстовый файл успешно конвертирован в {newFilePath}.";
            AddToConversionHistory(newFilePath, selectedFormat);
            _lastConvertedFilePath = newFilePath;
            LoadConversionHistory();
        }

        // Multimedia File Conversion Methods
        private void SelectMediaFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Media Files|*.mp3;*.wav;*.mp4;*.avi",
                Title = "Выберите мультимедийный файл"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedMediaFilePath = openFileDialog.FileName;
                SelectedMediaFilePath.Text = _selectedMediaFilePath;
                AnimateTextBlock(SelectedMediaFilePath);
            }
        }

        private async void ConvertMediaButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedMediaFilePath))
            {
                MediaStatusTextBlock.Text = "Пожалуйста, выберите мультимедийный файл.";
                return;
            }

            string selectedFormat = ((ComboBoxItem)MediaFormatComboBox.SelectedItem)?.Tag.ToString();
            if (string.IsNullOrEmpty(selectedFormat))
            {
                MediaStatusTextBlock.Text = "Пожалуйста, выберите формат.";
                return;
            }

            try
            {
                string newFilePath = Path.ChangeExtension(_selectedMediaFilePath, selectedFormat);
                await ConvertMediaFile(_selectedMediaFilePath, newFilePath);

                MediaStatusTextBlock.Text = $"Мультимедийный файл успешно конвертирован в {newFilePath}.";
                AddToConversionHistory(newFilePath, selectedFormat);
                _lastConvertedFilePath = newFilePath;
                LoadConversionHistory();
                OpenFolderButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MediaStatusTextBlock.Text = $"Ошибка: {ex.Message}";
            }
        }

        private async Task ConvertMediaFile(string inputFilePath, string outputFilePath)
        {
            var inputFile = new MediaFile { Filename = inputFilePath };
            var outputFile = new MediaFile { Filename = outputFilePath };

            using (var engine = new Engine())
            {
                await Task.Run(() => engine.Convert(inputFile, outputFile));
            }
        }

        // Database and History Methods
        private void AddToConversionHistory(string convertedPath, string targetFormat)
        {
            DatabaseHelper.AddHistoryItem(
                DateTime.Now,
                _selectedImagePath ?? _selectedTextFilePath ?? _selectedMediaFilePath,
                targetFormat,
                convertedPath
            );
        }

        private void LoadConversionHistory()
        {
            ConversionHistoryListBox.Items.Clear();
            var history = DatabaseHelper.GetHistory();
            foreach (var item in history)
            {
                ConversionHistoryListBox.Items.Add(
                    $"{item.Timestamp}: {item.OriginalPath} -> {item.TargetFormat} [{item.ConvertedPath}]"
                );
            }
        }

        private void MigrateOldHistory()
        {
            const string oldHistoryFilePath = "conversion_history.txt";
            if (File.Exists(oldHistoryFilePath))
            {
                var lines = File.ReadAllLines(oldHistoryFilePath);
                foreach (var line in lines)
                {
                    var parts = line.Split(new[] { ": ", " -> " }, 3, StringSplitOptions.None);
                    if (parts.Length == 3)
                    {
                        DatabaseHelper.AddHistoryItem(
                            DateTime.Parse(parts[0]),
                            parts[1],
                            parts[2],
                            "[old record]"
                        );
                    }
                }
                File.Delete(oldHistoryFilePath);
            }
        }

        // UI Helper Methods
        private void AnimateTextBlock(TextBlock textBlock)
        {
            textBlock.Opacity = 0;
            var fadeInAnimation = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.5));
            textBlock.BeginAnimation(UIElement.OpacityProperty, fadeInAnimation);
        }

        private BitmapEncoder GetEncoder(string format)
        {
            return format switch
            {
                "jpg" => new JpegBitmapEncoder(),
                "png" => new PngBitmapEncoder(),
                "bmp" => new BmpBitmapEncoder(),
                "gif" => new GifBitmapEncoder(),
                _ => null,
            };
        }

        // Event Handlers
        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastConvertedFilePath) && File.Exists(_lastConvertedFilePath))
            {
                string folderPath = Path.GetDirectoryName(_lastConvertedFilePath);
                Process.Start(new ProcessStartInfo("explorer.exe", folderPath));
            }
            else
            {
                MessageBox.Show("Нет последнего конвертированного файла.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void FormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            StatusTextBlock.Text = "Выбран формат: " + ((ComboBoxItem)FormatComboBox.SelectedItem)?.Content;
        }

        private void TextFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TextStatusTextBlock.Text = "Выбран формат: " + ((ComboBoxItem)TextFormatComboBox.SelectedItem)?.Content;
        }

        private void MediaFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MediaStatusTextBlock.Text = "Выбран формат: " + ((ComboBoxItem)MediaFormatComboBox.SelectedItem)?.Content;
        }

        private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            DatabaseHelper.ClearHistory();
            ConversionHistoryListBox.Items.Clear();
            HistoryStatusTextBlock.Text = "История конвертаций очищена.";
        }

        // Drag and Drop Handlers
        private void ImageTab_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                ShowPlusSign(PlusSignImage);
                ShowPlusSign(DragFilesText);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void ImageTab_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    _selectedImagePath = files[0];
                    SelectedImagePath.Text = _selectedImagePath;
                }
            }
            HidePlusSign(PlusSignImage);
            HidePlusSign(DragFilesText);
        }

        private void ImageTab_DragLeave(object sender, DragEventArgs e)
        {
            HidePlusSign(PlusSignImage);
            HidePlusSign(DragFilesText);
        }

        private void TextTab_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                ShowPlusSign(PlusSignText);
                ShowPlusSign(DragFilesTextText);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void TextTab_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    _selectedTextFilePath = files[0];
                    SelectedTextFilePath.Text = _selectedTextFilePath;
                }
            }
            HidePlusSign(PlusSignText);
            HidePlusSign(DragFilesTextText);
        }

        private void TextTab_DragLeave(object sender, DragEventArgs e)
        {
            HidePlusSign(PlusSignText);
            HidePlusSign(DragFilesTextText);
        }

        private void MediaTab_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                ShowPlusSign(PlusSignMedia);
                ShowPlusSign(DragFilesTextMedia);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void MediaTab_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    _selectedMediaFilePath = files[0];
                    SelectedMediaFilePath.Text = _selectedMediaFilePath;
                }
            }
            HidePlusSign(PlusSignMedia);
            HidePlusSign(DragFilesTextMedia);
        }

        private void MediaTab_DragLeave(object sender, DragEventArgs e)
        {
            HidePlusSign(PlusSignMedia);
            HidePlusSign(DragFilesTextMedia);
        }

        private void ShowPlusSign(TextBlock plusSign)
        {
            plusSign.Opacity = 1;
        }

        private void HidePlusSign(TextBlock plusSign)
        {
            plusSign.Opacity = 0;
        }
    }

    public class ConversionHistoryItem
    {
        public int Id { get; set; }
        public DateTime Timestamp { get; set; }
        public string OriginalPath { get; set; }
        public string TargetFormat { get; set; }
        public string ConvertedPath { get; set; }
    }

    public static class DatabaseHelper
    {
        private static string _databasePath = "conversionHistory.db";
        private static string _connectionString = $"Data Source={_databasePath};Version=3;";

        public static void InitializeDatabase()
        {
            if (!File.Exists(_databasePath))
            {
                SQLiteConnection.CreateFile(_databasePath);
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    string createTableQuery = @"
                        CREATE TABLE ConversionHistory (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Timestamp DATETIME NOT NULL,
                            OriginalPath TEXT NOT NULL,
                            TargetFormat TEXT NOT NULL,
                            ConvertedPath TEXT NOT NULL
                        )";
                    new SQLiteCommand(createTableQuery, connection).ExecuteNonQuery();
                }
            }
        }

        public static void AddHistoryItem(DateTime timestamp, string originalPath, string targetFormat, string convertedPath)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string insertQuery = @"
                    INSERT INTO ConversionHistory 
                    (Timestamp, OriginalPath, TargetFormat, ConvertedPath) 
                    VALUES (@timestamp, @originalPath, @targetFormat, @convertedPath)";

                var command = new SQLiteCommand(insertQuery, connection);
                command.Parameters.AddWithValue("@timestamp", timestamp);
                command.Parameters.AddWithValue("@originalPath", originalPath);
                command.Parameters.AddWithValue("@targetFormat", targetFormat);
                command.Parameters.AddWithValue("@convertedPath", convertedPath);
                command.ExecuteNonQuery();
            }
        }

        public static List<ConversionHistoryItem> GetHistory()
        {
            var history = new List<ConversionHistoryItem>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string selectQuery = "SELECT * FROM ConversionHistory ORDER BY Timestamp DESC";
                var command = new SQLiteCommand(selectQuery, connection);
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        history.Add(new ConversionHistoryItem
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            Timestamp = Convert.ToDateTime(reader["Timestamp"]),
                            OriginalPath = reader["OriginalPath"].ToString(),
                            TargetFormat = reader["TargetFormat"].ToString(),
                            ConvertedPath = reader["ConvertedPath"].ToString()
                        });
                    }
                }
            }
            return history;
        }

        public static void ClearHistory()
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                new SQLiteCommand("DELETE FROM ConversionHistory", connection).ExecuteNonQuery();
            }
        }
    }
}