using GoogleSheetsHelper;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//zdrojovy kod pro DLL C#
namespace GoogleSheet
{
    public class WritingToGoogleSheets
    {
        //musi byt static quli F#
        public static void WriteToGoogleSheets(string[] headers, string[][] rows, string jsonFileName, string id, string sheetName, int columnStart, int rowStart, bool complicatedCode)
        {
            var gsh = new GoogleSheetsHelper.GoogleSheetsHelper(jsonFileName, id);

            List<GoogleSheetCell> MyCells = new List<GoogleSheetCell>();
            List<GoogleSheetRow> MyRows = new List<GoogleSheetRow>();

            int numberOfColumns = headers.Length;
            int numberOfRows = 0;

            switch (complicatedCode)
            {
                case true:
                    numberOfRows = rows.ElementAt(0).Length; // takhle slozite, bo F# zrobilo vicerozmerove pole tak, jak ho zrobilo (csv)
                    break;
                case false:
                    numberOfRows = rows.Length;
                    break;
            }

            for (int i = -1; i < numberOfRows; i++)
            {
                if (i == -1)
                {
                    Enumerable.Range(0, numberOfColumns).ToList().ForEach(j => MyCells.Add(new GoogleSheetCell() { CellValue = headers[j].ToString(), IsBold = true }));
                    AddRows();
                }
                else
                {
                    switch (complicatedCode)
                    {
                        case true:
                            //Tady je 'j' a 'i' opacne (tj. v dll pro F# csv komplikovana varianta), nez mam v normalnim C# kodu
                            Enumerable.Range(0, numberOfColumns).ToList().ForEach(j => MyCells.Add(new GoogleSheetCell() { CellValue = rows[j][i].ToString() }));
                            break;
                        case false:
                            Enumerable.Range(0, numberOfColumns).ToList().ForEach(j => MyCells.Add(new GoogleSheetCell() { CellValue = rows[i][j].ToString() }));
                            break;
                    }
                    AddRows();
                }
            }

            gsh.AddCells(new GoogleSheetParameters() { SheetName = sheetName, RangeColumnStart = columnStart, RangeRowStart = rowStart }, MyRows);

            void AddRows()
            {
                GoogleSheetRow gsr = new GoogleSheetRow();

                gsr.Cells.AddRange(MyCells);
                MyRows.Add(gsr);
                MyCells.Clear();
            }
        }
    }
}
