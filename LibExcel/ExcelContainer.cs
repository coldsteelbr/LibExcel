using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.IO;

using nsExcel = Microsoft.Office.Interop.Excel;

namespace LibExcel
{
    /// <summary>
    /// Values - массив значений
    /// Считывает с первого листа
    /// </summary>
    public class ExcelReader
    {
        private string mWorkbookFileName;

        private nsExcel.Application mExcel = new nsExcel.Application();
        private nsExcel.Workbooks mWorkbooks;
        private nsExcel.Workbook mWorkbook;

        /// <summary>
        /// Ковертор нотации R1C1 в нотацию A1 (1024 => AMJ)
        /// </summary>
        /// <param name="p_R1C1">Число - номер столбца R1C1</param>
        /// <returns>Буквы - номер столбц</returns>
        public static string R1C1toA1(int p_R1C1) //1024 => AMJ
        {
            int res = p_R1C1;

            Collection<char> Letters = new Collection<char>();

            char currentChar;
            do
            {
                currentChar = (char)(res % 26);
                Letters.Add(currentChar);
                res /= 26;
            } while (res > 0);

            string result = "";
            foreach (char Letter in Letters)
            {
                result += ((char)(Letter + 96)).ToString();
            }
            char[] charResult = result.ToArray();
            Array.Reverse(charResult);
            return new string(charResult).ToUpper();
        }
        /// <summary>
        /// Конвертор нотации A1 в нотацию R1C1 (AMJ => 1024)
        /// </summary>
        /// <param name="p_A1">Буквы - номер столбца</param>
        /// <returns>Число - номер столбца в R1C1</returns>
        public static int A1toR1C1(string p_A1) //AMJ => 1024
        {
            char[] Letters = p_A1.ToLower().ToCharArray();
            for (int i = 0; i < Letters.Length; i++)
            {
                Letters[i] -= (char)96;
            }

            int result = (int)Letters[0];

            for (int i = 1; i < Letters.Length; i++)
            {
                result = result * 26 + Letters[i];
            }


            return result;// "R1C1";
        }

        public static object[,] GetOneBasedTwoDimenArray(int rows, int cols)
        {
            // 1-based 2-dimen array to be saved in Excel
            return (object[,])Array.CreateInstance(typeof(object), new int[] { rows, cols }, new int[] { 1, 1 });
        }

        public ExcelReader(string workbookFileName)
        {
            // Проверяем расширение файла
            if (Path.GetExtension(workbookFileName) == ".xlsx")
            {
                // Если указанный файл xlsx существует
                if (File.Exists(workbookFileName))
                {
                    mWorkbookFileName = workbookFileName;
                    mWorkbooks = mExcel.Workbooks;
                    mWorkbook = mWorkbooks.Open(mWorkbookFileName);
                } // TODO: else exception
            } // TODO: else exception

        }


        /// <summary>
        /// Получить значения конкретного листа по имени
        /// </summary>
        /// <param name="p_sheetName">Имя листа</param>
        /// <returns></returns>
        public object[,] GetSheetValues(string p_sheetName)
        {
            nsExcel.Worksheet currentSheet = (nsExcel.Worksheet)mWorkbook.Sheets[p_sheetName];
            nsExcel.Range currentRange = currentSheet.UsedRange;
            object[,] valuesToReturn = (object[,])currentRange.Value;
            return valuesToReturn;
        }

        /// <summary>
        /// Задать значения листа по имени
        /// </summary>
        /// <param name="p_sheetName">Имя листа</param>
        /// <param name="p_values">Массив значений</param>
        public void SetSheetValues(string p_sheetName, object[,] p_values)
        {
            nsExcel.Worksheet currentSheet = (nsExcel.Worksheet)mWorkbook.Sheets[p_sheetName];
            nsExcel.Range currentRange = currentSheet.Range[currentSheet.Cells[1, 1],
                   currentSheet.Cells[p_values.GetLength(0), p_values.GetLength(1)]];
            //currentRange.Value = p_values;
            currentRange.Formula = p_values;
        }

        /// <summary>
        /// закрывает файл, сохраняя изменения,  и закрывает Excel
        /// </summary>
        public void SaveChangesAndClose()
        {
            // saving the book
            mWorkbook.Save();
            // closing the book
            mWorkbook.Close();

            // realising com objects 
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(mWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(mWorkbooks);

            // closing and releasing the excel app 
            mExcel.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(mExcel);
        }

        /// <summary>
        /// Создаёт новый лист или возвращает имеющийся
        /// </summary>
        /// <param name="p"></param>
        public nsExcel.Worksheet CreateOrGetSheet(string p_SheetName)
        {
            nsExcel.Worksheet currentSheet;

            // получаем все имена листов
            Collection<string> sheetNames = new Collection<string>();
            foreach (nsExcel.Worksheet curSheet in this.mWorkbook.Worksheets)
            {
                sheetNames.Add(curSheet.Name);
            }

            // если наше имя содержится в списке
            if (sheetNames.Contains(p_SheetName))
            {
                // возвращаем имеющийся лист
                return (nsExcel.Worksheet)this.mWorkbook.Worksheets[p_SheetName];
            }
            else
            //иначе создаём новый
            {

                // создаём новый лист
                currentSheet = (nsExcel.Worksheet)mWorkbook.Worksheets.Add();
                // задаём имя
                currentSheet.Name = p_SheetName;
                // сохраняем книгу
                mWorkbook.Save();

                // возвращаем
                return currentSheet;
            }

        }
    }

}
