using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

using Word = Microsoft.Office.Interop.Word;

namespace BorDer_PR9
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        BorDer_PR9Entities db = new BorDer_PR9Entities();

        Книга book;
        int count = 1;

        public MainWindow()
        {
            InitializeComponent();

            //Создаем новый лист с только что созданными книгами
            var bookList = db.Книга.AsNoTracking().ToList();

            foreach (var book in bookList)
            {
                // Проверяем, что имя файла не пустое, и формируем полный путь
                if (!string.IsNullOrEmpty(book.Изображение))
                {
                    string cleanFileName = book.Изображение.Replace("\r", "").Replace("\n", "").Trim();
                    book.Изображение = System.IO.Path.Combine(AppContext.BaseDirectory, "Image", cleanFileName);
                }
            }

            listView.ItemsSource = bookList;
        }

        private void listView_SelectionChanged(object sender, SelectionChangedEventArgs e) //Выбор книги
        {
            try
            {
                var card = listView.SelectedItem as Книга;
                if (card != null)
                {
                    string fileName = card.Изображение.Replace("\r", "").Replace("\n", "").Trim();
                    string fullPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Image", fileName);
                    imgBook.Source = new BitmapImage(new Uri(fullPath));
                    textDescription.Text = card.Описание;
                    textPrice.Text = card.Цена.ToString();
                    gridDescription.Visibility = Visibility.Visible;

                    this.book = card;
                    count = 1;
                }
                else
                {
                    gridDescription.Visibility = Visibility.Collapsed;
                }
            }
            catch { }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).Name == "btnIncrease") // Нажат "+"
            {
                try
                {
                    count = int.Parse(tbCount.Text);
                    count++;
                    tbCount.Text = count.ToString();
                }
                catch { }
            }
            else //Нажат "-"
            {
                try
                {
                    count = int.Parse(tbCount.Text);
                    if (count <= 0)
                    {
                        throw new Exception("Нельзя уменьшить количество ниже 0!");
                    }
                    else
                    {
                        count--;
                        tbCount.Text = count.ToString();
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        private void tbCount_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (int.TryParse(tbCount.Text, out int count))
                {
                    if (count <= 0)
                    {
                        throw new Exception("Количество не может быть меньше или равно 0");
                    }
                }
                else
                {
                    throw new Exception("Ошибка конвертации числа");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                tbCount.Text = 1.ToString();
            }
        }

        private void createOrder_Click(object sender, RoutedEventArgs e)
        {
            createCheck();
        }

        public void createCheck()
        {
            try
            {

                var Продажа = new Продажи()
                {
                    Дата = DateTime.Now,
                    Книга = book.Код,
                    Количество = count,
                    Итог = book.Цена * count
                };
                db.Продажи.Add(Продажа);

                
                    db.SaveChanges();
                
                MessageBox.Show($"Продажа сохранена! Присвоен номер чека: {Продажа.НомерЧека}");


                var Word = new Word.Application();
                Word.Visible = false;
                var document = Word.Documents.Open(Environment.CurrentDirectory + @"\Template.docx");

                replaceWord("{НомерЧека}", (Продажа.НомерЧека).ToString(), document);
                replaceWord("{Дата}", (Продажа.Дата).ToString(), document);
                replaceWord("{Название}", (book.Название).ToString().Trim(), document);
                replaceWord("{Количество}", (Продажа.Количество).ToString(), document);
                replaceWord("{ИтоговаяСумма}", (Продажа.Итог).ToString(), document);

                document.SaveAs2(System.IO.Path.Combine(AppContext.BaseDirectory, "Check_" + Продажа.НомерЧека));
                MessageBox.Show("Билет сохранен");

            }
            catch (Exception e) { MessageBox.Show(e.Message); }

        }

        private void replaceWord(string subToReplace, string text, Word.Document document)
        {
            var range = document.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: subToReplace, ReplaceWith: text);
        }


    }
}
