using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace ConsoleWord
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application(); //Создаем экземпляр приложения
            Document doc = app.Documents.Add(Visible:true); //Создание видимого приложения
                Range r = doc.Range();
                r.Text = "Hello Word";
                                                               //r.Bold = 20;
                Table t = doc.Tables.Add(r,10,7); // was 5, 5
                t.Borders.Enable = 1; //Видимые границы таблицы
                foreach (Row row in t.Rows)
                {
                    foreach(Cell cell in row.Cells)
                    {
                        if(cell.RowIndex == 1)
                        {
                            cell.Range.Text = "Колонка" + cell.ColumnIndex.ToString(); //Текст колонок
                            cell.Range.Bold = 1;
                            cell.Range.Font.Name = "verdana"; //Шрифт
                            cell.Range.Font.Size = 10; //Размер

                            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; //Выравнивание
                            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; //Выравнивание
                        }
                        else{
                            //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                            cell.Range.Text = "Hello World"; //Заполнение
                        }
                    }
                }

                doc.Save(); //Сохраним документ
            app.Documents.Open(@"Doc2.docx"); //Откроем документ в папке с программой
            Console.ReadKey(); //Ожидаем ввод
            try //попытка закрыть
            {
                doc.Close(); //Закрыли документ
                app.Quit(); //Закрыли приложение
            }
            catch(Exception e) //Отлов ошибки
            {
                Console.WriteLine(e.Message); //Вывод ошибки
            }
            Console.ReadKey();
        }
    }
}
