// This is an independent project of an individual developer. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++ and C#: http://www.viva64.com

using System.Collections.Generic;
using System.Linq;
using System.Windows.Documents;

namespace DocxToWpf
{
    /// <summary> Класс, преобразующий RowSpan ячеек таблицы из word в WPF. </summary>
    public class RowSpanConverter
    {
        /// <summary> Конструктор. </summary>
        /// <param name="table"> Таблица, которая будет обрабатываться. </param>
        public RowSpanConverter(Table table)
        {
            _table = table;
        }

        /// <summary> Добавление строки таблицы. </summary>
        public void AddRow()
        {
            _grid.Add(new List<(int row, int column)>());
        }

        /// <summary> Добавление ячейки в конец последней строки. </summary>
        /// <param name="cellColumnSpan"> Размер ячейки по горизонтали. </param>
        /// <param name="cellHasContinue"> Является ли эта ячейка продолжением предыдущей по вертикали. </param>
        public void AddCell(int cellColumnSpan, bool cellHasContinue)
        {
            // Координаты новой ячейки в таблице.
            int tableRow;
            int tableColumn;
            
            if (cellHasContinue)
            {
                // Координаты новой ячейки в сетке.
                int currentRow = _grid.Count - 1;
                int currentColumn = _grid.Last().Count;
                
                (tableRow, tableColumn) = _grid[currentRow - 1][currentColumn];
            }
            else
            {
                tableRow = _table.RowGroups[0].Rows.Count - 1;
                tableColumn = _table.RowGroups[0].Rows.Last().Cells.Count - 1;
            }

            for (int i = 0; i < cellColumnSpan; i++)
            {
                _grid.Last().Add((tableRow, tableColumn));
            }
        }

        /// <summary> Преобразование таблицы к правильному виду. </summary>
        public void SetupTable()
        {
            List<(int row, int column)> cellsToDelete = new List<(int row, int column)>();
            
            for (int i = 1; i < _grid.Count; i++)
            {
                for (int j = 0; j < _grid[i].Count; j++)
                {
                    // Проходим по всем ячейкам сетки.
                    
                    // Координаты текущей ячейки таблицы
                    (int currentRow, int currentColumn) = GridToTableWithoutRowSpan(i, j);
                    
                    // Координаты, которые должны быть у текущей ячейки таблицы.
                    (int trueRow, int trueColumn) = _grid[i][j];
                    
                    // Если ячейка является продолжением предыдущей по горизонтали, пропускаем. 
                    if (j >= 1 && currentColumn == GridToTableWithoutRowSpan(i, j-1).column) continue;

                    if (currentRow == trueRow) continue;
                    
                    TableCell trueCell = _table.RowGroups[0].Rows[trueRow].Cells[trueColumn];
                    TableCell currentCell = _table.RowGroups[0].Rows[currentRow].Cells[currentColumn];

                    // Записываем данные из текущей ячейки в правильную.
                    // Цикл нужен, т.к. trueCell.Blocks.AddRange(currentCell.Blocks) не работает. 
                    while (currentCell.Blocks.Count > 0)
                    {
                        trueCell.Blocks.Add(currentCell.Blocks.FirstBlock);
                        currentCell.Blocks.Remove(currentCell.Blocks.FirstBlock);
                    }

                    // Помечаем неправильные ячейки на удаление.
                    cellsToDelete.Add((currentRow, currentColumn));
                    trueCell.RowSpan++;
                }
            }

            // Удаляем помеченные ячейки.
            for (int i = cellsToDelete.Count - 1; i >= 0; i--)
            {
                (int row, int column) = cellsToDelete[i];
                _table.RowGroups[0].Rows[row].Cells.RemoveAt(column);
            }
        }

        #region Private definitions

        /// <summary> Таблица, которая будет обрабатываться. </summary>
        private readonly Table _table;
        
        /// <summary> Сетка, в ячейках которой указаны координаты соответствующей ячейки таблицы. </summary>
        private readonly List<List<(int row, int column)>> _grid = new List<List<(int row, int column)>>();

        /// <summary> Преобразует координаты сетки в координаты таблицы без учета RowSpan. </summary>
        /// <param name="row"> Строка сетки. </param>
        /// <param name="column"> Столбец сетки. </param>
        /// <returns></returns>
        private (int row, int column) GridToTableWithoutRowSpan(int row, int column)
        {
            TableRow tableRow = _table.RowGroups[0].Rows[row];
            int tableColumn = 0;
            int gridColumn = 0;
            
            foreach (TableCell cell in tableRow.Cells)
            {
                gridColumn += cell.ColumnSpan;
                if (gridColumn > column) break;
                tableColumn++;
            }

            return (row, tableColumn);
        }
        
        #endregion
    }
}