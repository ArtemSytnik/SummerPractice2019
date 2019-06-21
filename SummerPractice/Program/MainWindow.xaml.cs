using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Program
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Диалог откоытия файла
        OpenFileDialog openDialog = new OpenFileDialog();
        //Диалог сохранения файла
        SaveFileDialog saveDialog = new SaveFileDialog();
        //Результат выполнения исследований
        double result = -1;

        public MainWindow()
        {
            InitializeComponent();
            //Фильтры диалогов
            openDialog.Filter = "Json files (*.json)|*.json";
            saveDialog.Filter = "Excel Worksheets|*.xlsx";
        }
        //Обработка данных
        private void Calculating_Click(object sender, RoutedEventArgs e)
        {
            if (openDialog.ShowDialog() != false)
            {
                try
                {
                    //Создание файлового потока
                    FileStream fs = new FileStream(openDialog.FileName, FileMode.Open, FileAccess.Read);
                    using (fs)
                    {
                        //Чтение файла
                        string text = File.ReadAllText(openDialog.FileName, Encoding.GetEncoding(1251));
                        //Приведение файла к стандарту Json
                        text = text.Replace(",\"items\":\n", String.Empty);
                        text = text.Replace("}{", "},{");
                        text = text.Replace("]}", "]");
                        //Выполнение исследований
                        result = percentProcessing(text);
                        //Вывод результатов
                        result_visual.Content = $"Результат : {result} %";
                    }
                }
                catch(Exception exc)
                {
                    MessageBox.Show("Ошибка:\n" + exc.Message, "Ошибка выполнения", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        //Десеривализация и выполнение исследований
        double percentProcessing(string text)
        {
            //Десериализация
            var studentMarks = JsonConvert.DeserializeObject<List<Student>>(text).GroupBy(f => f.id).Select(grp => grp.ToList()).ToList();
            //% студентов, которые учаться только на оценки «4 и 5»
            double value = 0;
            //общее колличество студентов
            double allStudents=0;
            //колличество студентов подходящих под условие
            double requiredStudents =0;
            //оценка студента
            int mark;
            //проверка подходит ли студент под условие
            bool error = false;
            foreach(List<Student> x in studentMarks)
            {
                allStudents++;
                foreach(Student element in x)
                {
                    if (!error)
                    {
                        //проверка условия
                        if (int.TryParse(element.name, out mark) == true)
                        {
                            if (mark < 75)
                                error = true;
                        }
                        else
                        {
                            error = true;
                        }
                    }
                }
                if (!error)
                    //студент подходит под условие
                    requiredStudents++;
                error = false;
            }
            //расчет процента студентов подходящих под условие
            value = requiredStudents / allStudents * 100;
            value = Math.Round(value, 3);
            return value;
        }
        //Сохранение результатов исследований
        private void Saving_Click(object sender, RoutedEventArgs e)
        {
            //Проверка было ли произведено исследование
            if(result == -1)
            {
                MessageBox.Show("Вычисление процента студентов не было произведено.", "Ошибка выполнения", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (saveDialog.ShowDialog() != false)
            {
                try
                {
                    //Создание файла содержащего результаты исследований
                    Excel.Application ObjExcel = new Excel.Application();
                    //Книга
                    Excel.Workbook ObjWorkBook;
                    //Таблица
                    Excel.Worksheet ObjWorkSheet;
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                    //Заполнение значениями
                    ObjWorkSheet.Name = "Результат исследований";
                    ObjWorkSheet.Cells[1, 1] = "% студентов, которые учаться только на оценки «4 и 5»";
                    ObjWorkSheet.Cells[1, 2] = result;
                    ObjWorkSheet.UsedRange.Columns.AutoFit();
                    //Сохранение
                    ObjWorkBook.SaveAs(saveDialog.FileName);
                    ObjExcel.Quit();
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Ошибка:\n" + exc.Message, "Ошибка выполнения", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

    }
}
