using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace ExcelMapWord
{
    internal class ExcelHelper
    {
        private string _excelPath;
        private string _sheetName;
        private int _rowStart;
        private int _rowEnd;
        private int _rowCount;
        //private string _rowRange;
        private string _wordSamplePath;
        private string _saveDir;

        #region 构造函数
        public ExcelHelper(CmdInputArgument inputArgument)
        {
            if (File.Exists(inputArgument.ExcelPath))
            {
                _excelPath = inputArgument.ExcelPath;
            }
            else
            {
                throw new FileNotFoundException("文件不存在", _excelPath);
            }

            if (string.IsNullOrEmpty(inputArgument.SheetName))
            {
                throw new ArgumentNullException(nameof(inputArgument.SheetName));
            }
            else
            {
                _sheetName = inputArgument.SheetName;
            }

            #region row
            if (string.IsNullOrEmpty(inputArgument.Row))
            {
                throw new ArgumentNullException(nameof(inputArgument.Row));
            }
            else
            {
                var rowRange = inputArgument.Row.Split('-');
                if (rowRange.Length == 1)
                {
                    _rowCount = 1;
                    _rowStart = int.Parse(rowRange[0]);
                }
                else
                {
                    _rowStart = int.Parse(rowRange[0]);
                    _rowEnd = int.Parse(rowRange[1]);
                    _rowCount = int.Parse(rowRange[1].Trim()) - int.Parse(rowRange[0].Trim()) + 1;

                    if (_rowStart > _rowEnd)
                    {
                        throw new Exception("Error: rowStart > rowEnd");
                    }
                }
            }
            #endregion

            if (File.Exists(inputArgument.WordSamplePath))
            {
                _wordSamplePath = inputArgument.WordSamplePath;
            }
            else
            {
                throw new ArgumentNullException(nameof(inputArgument.WordSamplePath));
            }

            _saveDir = Path.Combine(Path.GetDirectoryName(_excelPath), "output");
        }
        #endregion

        #region + public void StartMapToWord()
        public void StartMapToWord()
        {
            FileStream fs = null;
            XLWorkbook wb = null;

            try
            {
                byte[] sample_buffer = File.ReadAllBytes(_wordSamplePath);
                
                using (fs = new FileStream(_excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (wb = new XLWorkbook(fs))
                {
                    IXLWorksheet sheet = wb.Worksheet(_sheetName);
                    IXLRows rows = sheet.Rows(_rowStart, _rowEnd);
                    foreach (var row in rows)
                    {
                        //each row map to word
                        try
                        {
                            RowMapWord(row, sample_buffer);
                        }
                        catch (Exception ex)
                        {
                            LogHelper.Log($"Map错误: {ex.Message}\r\n行号: {row.RowNumber()}\r\n");
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                wb?.Dispose();
                fs?.Dispose();
            }

        }
        #endregion

        #region + private void RowMapWord(IXLRow row, byte[] sample_buffer)
        private void RowMapWord(IXLRow row, byte[] sample_buffer)
        {
            DocX doc = null;
            MemoryStream ms = new MemoryStream();

            try
            {
                ms.Write(sample_buffer, 0, sample_buffer.Length);
                ms.Seek(0, SeekOrigin.Begin);

                doc = DocX.Load(ms);

                foreach (var cell in row.CellsUsed())
                {
                    StringReplaceTextOptions options = new StringReplaceTextOptions
                    {
                        SearchValue = $"[{cell.WorksheetColumn().ColumnLetter()}]",
                        NewValue = cell.Value.ToString()
                    };

                    doc.ReplaceText(options);
                }

                string saveFileName = Path.Combine(_saveDir, row.FirstCellUsed().Value.ToString() + "_" + row.RowNumber());
                doc.SaveAs(saveFileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                doc?.Dispose();
                ms?.Dispose();
            }
        }
        #endregion
    }
}
