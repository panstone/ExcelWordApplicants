# ExcelWordApplicants
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Forward
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void pusi(Excel.Worksheet currentSheet, Excel.Application excelApp)
        {
            //выполняем поиск максимально кол во строк колво столбцов
            int MaxRow;
            int MaxCol;
            string str;
            string dat;
            string dattemp = "";
            Object bul;
            Object ofv;
            DateTime dtfv;


            double kol_list;
            double kol_ekz;

            string strtemp = "";
            int y = 0;//для получения документов
            string[] myArr = new string[20];//массив для записи имени доков
            string vall = "SHABLON";//первый проход для правильного отображения названий созданных документов(ВАЖНО!!!)
            excelApp.Workbooks.Open(tbExcel.Text);
            currentSheet = (Excel.Worksheet)excelApp.Workbooks[1].Worksheets[1];
            MaxRow = currentSheet.UsedRange.Rows.Count;
            MaxCol = currentSheet.UsedRange.Columns.Count;
            Excel.Range range;
            range = currentSheet.UsedRange;
            //по строке как в матрице
            string znach = "tot";
            DateTime date1 = new DateTime(2008, 3, 1, 7, 0, 0);
            // string date = "";
            DateTime date2 = new DateTime(2008, 3, 1, 7, 0, 0);
            //открываем ворд
            Word.Application word = new Word.Application();
            for (int cCnt = 2; cCnt <= MaxRow; cCnt++)//по строкам начиная со второй
            {//!!!!!!!!!!!!!!сделать обработчик выведет сообщение о том что вся таблица пустая!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                
                    str = (string)(range.Cells[cCnt, 6] as Excel.Range).Value2;//получаем локально данные 6 столбца
                                                                               //
                   
                    if (str != null && strtemp != str)
                    //получили первый маил
                    {
                    


                        Word.Document doc = word.Documents.Add(tbWord.Text);


                        //пошли по строке пересчитывая столбцы 
                        for (int cCft = 1; cCft <= MaxCol; cCft++)
                        {

                            //первый проход проверяет первую ячейку ниже код   
                            
                            //еще надо обернуть это все в проверку нулевой ссылки
                            if ((range.Cells[cCnt, cCft] as Excel.Range).Value2 != null)
                            {   

                                // если не пустая ячейка, тогда все тут крутиться
                                if ((range.Cells[cCnt, cCft] as Excel.Range).Value2.GetTypeCode() == TypeCode.String)//если попался доубля то он идет вниз
                                {
                                    dat = (string)(range.Cells[cCnt, cCft] as Excel.Range).Value2;
                                    switch (cCft)
                                    {
                                        case 2:
                                            doc.Bookmarks["ZAKL"].Range.Text = dat;//номер регистрации
                                            break;
                                        case 3:
                                            doc.Bookmarks["UCHREGD"].Range.Text = dat;//Учетно рег действие
                                            break;
                                        case 5:
                                            doc.Bookmarks["ZAYAVIL"].Range.Text = dat;
                                            doc.Bookmarks["ZAYAVIL2"].Range.Text = dat;//фИО заявителя
                                            break;
                                        case 6:
                                            doc.Bookmarks["EMAIL"].Range.Text = dat;//электронный адрес заявителя
                                            break;
                                        case 12://готовит докупенты с листами копиями и т.д.
                                            y = 1;
                                            dattemp = "";
                                            for (int cCnt2 = 2; cCnt2 <= MaxRow; cCnt2++)
                                                if ((string)(range.Cells[cCnt2, 6] as Excel.Range).Value2 == str && (string)(range.Cells[cCnt2, 12] as Excel.Range).Value2 != null)
                                                {
                                                    dattemp += y.ToString() + ". " + (string)(range.Cells[cCnt2, 12] as Excel.Range).Value2 + " ";

                                                    if ((range.Cells[cCnt2, 16] as Excel.Range).Value2 != null)
                                                    {
                                                        kol_list = (double)(range.Cells[cCnt2, 16] as Excel.Range).Value2;
                                                        dattemp += kol_list.ToString() + " л. ";
                                                    }
                                                    if ((range.Cells[cCnt2, 17] as Excel.Range).Value2 != null)
                                                    {
                                                        kol_ekz = (double)(range.Cells[cCnt2, 17] as Excel.Range).Value2;
                                                        dattemp += kol_ekz.ToString() + " экз. ";
                                                    }

                                                    // kol_list = kol_list_obj.ToString();
                                                    //    if (kol_list != null)


                                                    // if ((string)(range.Cells[cCnt2, 16] as Excel.Range).Value2 != null) { dattemp += (string)(range.Cells[cCnt2, 16] as Excel.Range).Value2 + " л "; }
                                                    // if ((string)(range.Cells[cCnt2, 17] as Excel.Range).Value2 != null) { dattemp += (string)(range.Cells[cCnt2, 17] as Excel.Range).Value2 + " экз "; }
                                                    // if ((string)(range.Cells[cCnt2, 18] as Excel.Range).Value2 != null) { dattemp += (string)(range.Cells[cCnt2, 18] as Excel.Range).Value2 + " кол. копий "; }
                                                    dattemp += " \r\n ";
                                                    y++;
                                                }
                                            doc.Bookmarks["NAIMDOC1"].Range.Text = dattemp;
                                            break;
                                        case 20:
                                            doc.Bookmarks["KN"].Range.Text = dat;//кадастровый номер
                                            break;
                                        case 21:
                                            doc.Bookmarks["ADRESS"].Range.Text = dat;//адрес объекта
                                            break;
                                        /*  case 22:
                                              doc.Bookmarks["PLANDOG"].Range.Text = dat;//дата окончания
                                              break;*/
                                        case 23:
                                            doc.Bookmarks["TYPEOBJ"].Range.Text = dat;//тип объекта
                                            break;
                                        default:
                                            doc.Bookmarks["NAIMDOC1"].Range.Text = myArr[0];
                                            break;
                                    }
                                    //номер регистрации 2 столбец
                                    /* doc.Bookmarks["DATAREG"].Range.Text = dat;//принято в работу 1ый столбец
                                     doc.Bookmarks["UCHREGD"].Range.Text = dat;//учетно регистрационные действия 3 столбец их перед добавлением надо сравнивать
                                     //4 столбец это способ получения документа
                                     doc.Bookmarks["ZAYAVIL"].Range.Text = dat;//заявитель 5 столбец(возможно нужно будет редактировать)
                                     doc.Bookmarks["EMAIL"].Range.Text = dat;//Имаил это 6 столбец тоже вставляем
                                     //7 столбец это номер телефона 
                                     //8 нахер не нужен это адрес для почты заяввителя
                                     //9 столбец представитель 
                                     //10 столбец электронная почта представителя нигде не указывается
                                     //11 столбец адрес почты представителя
                                     doc.Bookmarks["NAIMDOC1"].Range.Text = dat;//12 столбец наименование документа но тут вопрос их несколькодва или 3 надо смотреть
                                     //13 серия номер документа не нужны
                                     //14 номер просто номер 
                                     doc.Bookmarks["DATAVID1"].Range.Text = dat;//15 дата выдачи первый документ их два или три надо смотреть
                                     doc.Bookmarks["KOLVOLIST1"].Range.Text = dat;//16 количество листов так же как и др их может 2 или 3 быть
                                     doc.Bookmarks["KOLVOEKZ1"].Range.Text = dat;//17 количестно экземпларов так же как и др их может 2 или 3 быть
                                     //18 количество копий  
                                     //19 количество 
                                     doc.Bookmarks["KN"].Range.Text = dat;//20 столбец кадастровый номер
                                     doc.Bookmarks["ADRESS"].Range.Text = dat;//21 столбец там  адрес он пихается в ворд(тоже надо редактировать
                                     doc.Bookmarks["PLANDOG"].Range.Text = dat;//22 столбец срок завершения тоже в ворде
                                     doc.Bookmarks["TYPEOBJ"].Range.Text = dat;//23 столбец тип объекта типа земельный участок */
                                }//если вдруг ошибся с типом то продолжает дальше по циклу

                                else if ((range.Cells[cCnt, cCft] as Excel.Range).Value2.GetTypeCode() == TypeCode.Double)
                                {
                                    //данные в первом столбце
                                    //если этот объект возвращается как дата
                                    try
                                    {
                                        ofv = (object)(range.Cells[cCnt, cCft] as Excel.Range).Value2;
                                        dtfv = DateTime.FromOADate((double)ofv);// передает дату

                                        switch (cCft)
                                        {
                                            case 1:
                                                date1 = dtfv;//заданную перезаписываем на полученную
                                                doc.Bookmarks["SECONDNAME"].Range.Text = date1.ToString();
                                                doc.Bookmarks["SECONDNAME2"].Range.Text = date1.ToString();//номер регистрации
                                                                                                           // }
                                                break;
                                            case 22:
                                                date2 = dtfv;//заданную перезаписываем на полученную
                                                doc.Bookmarks["PLANDOG"].Range.Text = date2.ToString();
                                                break;
                                                // MessageBox.Show(dtfv.ToString());//а вот и сообщуха
                                        }

                                    }
                                    catch
                                    {
                                        bul = (object)(range.Cells[cCnt, cCft] as Excel.Range).Value2;
                                    }
                                }
                                doc.Bookmarks["NAIMDOC1"].Range.Text = myArr[0];
                            }
                            else
                            {
                                //можно создать переменную и пусть выполняет действие
                                //  MessageBox.Show("Попала пустая ячейка");//возвращает если равно нулю значение ячейки
                            }
                            //затем он дрлжен сохранить документ ворд как закончит и присупить к открытию нового и сохранения туда данных со следующим мылом 
                            vall = str;
                            doc.SaveAs(FileName: @"C:\Users\Панаскин ДН\Documents\" + vall + ".docx");//путь рабочий
                            strtemp = str;
                        }
                        doc.Close();
                        //теперь самое интересное  
                        if (str != znach)
                        {
                            znach = str;
                            listBox1.Items.Add(znach);//получили в листбокс все имаилы адресами                           
                        }        
                }
            }
        } 
        private void OpenExcel_Click(object sender, EventArgs e)
        {
            // Отображает диалоговое окно openfiledialog, так что пользователь может выбрать файл Excel
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog1.Title = "Выберите файл Excel";
            // Показать диалоговое окно.
            // Если пользователь нажал ОК в диалоговом окне и
            // а .xls файл был выбран, откройте его.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Назначить Excel в потоке в собственность в Excel формы.
                //this.Cursor = new Cursor(openFileDialog1.OpenFile());
                //передает путь в текстбокс
                tbExcel.Text = openFileDialog1.FileName;//получили ссылку
            }
        }
        private void result_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp;
            Excel.Worksheet currentSheet;
            // Excel.Workbook WorkBook;
            try
            {
                //открытие excel
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.Workbooks.Open(tbExcel.Text);//получили рабочую книгу можно 
                currentSheet = (Excel.Worksheet)excelApp.Workbooks[1].Worksheets[1];
                // WorkBook = excelApp.Workbooks.Open(tbExcel.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            }
            catch (Exception)
            {
                MessageBox.Show("Excel не исправен");
                return;
            }
            pusi(currentSheet, excelApp);//в этом месте прога получет максимальное количество строк и столбцов
            MessageBox.Show("Программа завершила свою работу!");
            excelApp.Quit();
        }
        private void OpenWord_Click_1(object sender, EventArgs e)
        {
            // Отображает диалоговое окно openfiledialog, так что пользователь может выбрать файл Excel
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Word Files|*.doc;*.docx";
            openFileDialog2.Title = "Выберите файл Word";
            // Показать диалоговое окно.
            // Если пользователь нажал ОК в диалоговом окне и
            // а .xls файл был выбран, откройте его.
            if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Назначить Excel в потоке в собственность в Excel формы.
                //this.Cursor = new Cursor(openFileDialog1.OpenFile());
                //передает путь в текстбокс
                tbWord.Text = openFileDialog2.FileName;//получили ссылку
            }
        }
    }
}
