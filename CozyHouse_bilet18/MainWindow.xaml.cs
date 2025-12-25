using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CozyHouse_bilet18
{
    public partial class MainWindow : Window
    {
        УютныйДомEntities db = new УютныйДомEntities();

        double totalSum = 0;
        Материал selectedMaterial;
        double height = 0;
        double width = 0;

        public MainWindow()
        {
            InitializeComponent();
            LoadMaterials();
        }

        private void LoadMaterials()
        {
            typeMaterialCombo.ItemsSource = db.Материал.ToList();
            typeMaterialCombo.DisplayMemberPath = "Наименование";
        }

        // РАСЧЁТ СТОИМОСТИ
        public bool CalculateCost()
        {
            if (typeMaterialCombo.SelectedItem == null ||
                string.IsNullOrWhiteSpace(widthText.Text) ||
                string.IsNullOrWhiteSpace(heightText.Text))
            {
                MessageBox.Show("Выберите материал и заполните размеры");
                return false;
            }

            selectedMaterial = (Материал)typeMaterialCombo.SelectedItem;

            try
            {
                height = double.Parse(heightText.Text.Replace('.', ','));
                width = double.Parse(widthText.Text.Replace('.', ','));
            }
            catch
            {
                MessageBox.Show("Ошибка формата размеров");
                return false;
            }

            if (selectedMaterial.ЦенаЗаКвМетр == null)
            {
                MessageBox.Show("У материала не указана цена");
                return false;
            }

            totalSum = height * width * selectedMaterial.ЦенаЗаКвМетр.Value;

            infMaterialLab.Content =
                $"Размер: {height:F2} x {width:F2}\n" +
                $"Материал: {selectedMaterial.Наименование}\n" +
                $"Стоимость: {totalSum:F2} руб.";

            infMaterialLab.Visibility = Visibility.Visible;
            return true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CalculateCost();
        }

        // ОФОРМЛЕНИЕ ПОКУПКИ И ЧЕКА
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (!CalculateCost())
                return;

            // сохраняем покупку
            Покупка purchase = new Покупка
            {
                ДатаПокупки = DateTime.Now
            };

            db.Покупка.Add(purchase);
            db.SaveChanges();

            // состав покупки
            СоставПокупки item = new СоставПокупки
            {
                Покупка = purchase.Код,
                Материал = selectedMaterial.Код,
                Длина = height,
                Ширина = width,
                Сумма = totalSum
            };

            db.СоставПокупки.Add(item);
            db.SaveChanges();

            // создаём чек
            CreateReceiptWord(purchase);
        }


        // СОЗДАНИЕ ЧЕКА WORD 
        private void CreateReceiptWord(Покупка purchase)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = word.Documents.Add();

            word.Visible = false;

            // Общий шрифт
            doc.Content.Font.Name = "Times New Roman";
            doc.Content.Font.Size = 10;

            void AddText(string text)
            {
                var p = doc.Content.Paragraphs.Add();
                p.Range.Text = text;
                p.Range.InsertParagraphAfter();
            }

            // ШАПКА
            AddText("ООО \"Уютный Дом\"");
            AddText("Добро пожаловать");
            AddText("ККМ 00075411    #3969");
            AddText("ИНН 1087746942040");
            AddText("ЭКЛЗ 3851495566");
            AddText($"Чек №{purchase.Код}");
            AddText($"{DateTime.Now:dd.MM.yyyy HH:mm} СИС");
            AddText("");

            // ТАБЛИЦА 
            var table = doc.Tables.Add(doc.Range(doc.Content.End - 1), 6, 2);
            table.Borders.Enable = 0;

            table.Columns[1].Width = 120;
            table.Columns[2].Width = 100;

            table.Cell(1, 1).Range.Text = "наименование товара";
            table.Cell(1, 1).Merge(table.Cell(1, 2));
            table.Cell(1, 1).Range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;


            table.Cell(2, 1).Range.Text = "жалюзи";
            table.Cell(2, 2).Range.Text = $"{height:F2} x {width:F2}";
            table.Cell(2, 2).Range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;


            table.Cell(3, 1).Range.Text = "материал";
            table.Cell(3, 2).Range.Text = selectedMaterial.Наименование;

            string sum = totalSum.ToString("F2");

            table.Cell(4, 1).Range.Text = "Итог";
            table.Cell(4, 2).Range.Text = sum;

            table.Cell(5, 1).Range.Text = "Сдача";
            table.Cell(5, 2).Range.Text = "0";

            table.Cell(6, 1).Range.Text = "Сумма итого:";
            table.Cell(6, 2).Range.Text = sum;

            // НИЗ ЧЕКА
            AddText("");
            AddText("***********************");
            AddText("00003751# 059705");

            // СОХРАНЕНИЕ
            string fileName = $"Чек_{purchase.ДатаПокупки:yyyyMMdd}_{totalSum:F2}.docx";
            string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);


            doc.SaveAs2(path);
            doc.Close(false);
            word.Quit();

            MessageBox.Show("Чек создан");

            try
            {
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                MessageBox.Show("Не удалось открыть файл");
            }
        }

        // ВВОД ТОЛЬКО ЧИСЕЛ 
        private void DoubleNum_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            string newText = tb.Text.Insert(tb.SelectionStart, e.Text);
            e.Handled = !Regex.IsMatch(newText, @"^\d{1,4}([.,]\d{0,2})?$");
        }
    }
}
