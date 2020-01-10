using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApplication1
{
     class OperateExcel
    {
        String filePath;
        IWorkbook workBook;
        ISheet workSheet;
        FileStream excelStream;

        public OperateExcel(String filePath) {
            this.filePath = filePath;
            if (!File.Exists(filePath))
            {
                if (filePath.IndexOf(".xlsx") > 0)
                {
                    workBook = new XSSFWorkbook();
                }
                else if (filePath.IndexOf(".xls") > 0)
                {
                    workBook = new HSSFWorkbook();
                }
                excelStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            }
            else
            {
                excelStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);
                if (filePath.IndexOf(".xlsx") > 0)
                {
                    workBook = new XSSFWorkbook(excelStream);
                }
                else if (filePath.IndexOf(".xls") > 0)
                {
                    workBook = new HSSFWorkbook(excelStream);
                }
            }
         }

        public ISheet CreateOrOpenWorkSheet(String sheetName) {
            if (this.workBook.GetSheet(sheetName) == null)
            {
                this.workBook.CreateSheet(sheetName);
            }
            workSheet = this.workBook.GetSheet(sheetName);
            return workSheet;
         }

        public void SerializableBaseDatas(List<dynamic> baseDatas) {
            string sheetName = this.workSheet.SheetName;
            this.workBook.RemoveSheetAt(this.workBook.GetSheetIndex(sheetName));
            this.workSheet=this.workBook.CreateSheet(sheetName);
            IRow row=this.workSheet.CreateRow(0);
            int columnIndex = 0;
            foreach (var key in baseDatas[0].GetProperty())
            {
                row.CreateCell(columnIndex).SetCellValue(key);
                columnIndex++;
            }
            int rowIndex = 1;
            foreach (var baseData in baseDatas)
            {
                IRow dataRow = this.workSheet.CreateRow(rowIndex);
                columnIndex = 0;
                foreach (var value in baseData.GetValue())
                {
                    dataRow.CreateCell(columnIndex).SetCellValue(value);
                    columnIndex++;
                }
                rowIndex++;
            }
            
        }

        public List<dynamic> DeSerialableBaseDatas(List<String> keys) {
            List<dynamic> BaseDatas = new List<dynamic>();
            List<int> columnIndexs = new List<int>();
            foreach (var key in keys)
            {
                int colunnIndex = -1;
                for (int i = 0; i < this.workSheet.GetRow(0).LastCellNum; i++)
                {
                    if (this.workSheet.GetRow(0).GetCell(i).StringCellValue == key)
                    {
                        colunnIndex = i;
                        break;
                    }
                }
                columnIndexs.Add(colunnIndex);
            }
            for (int i = 0; i < this.workSheet.LastRowNum; i++)
            {
                dynamic BaseData = new DynamicModel();
                for (int j = 0; j < keys.Count; j++)
                {
                    BaseData.PropertyName = keys[j];
                    BaseData.Property = this.workSheet.GetRow(i).GetCell(columnIndexs[j]).StringCellValue;
                }
                BaseDatas.Add(BaseData);
            }
            return BaseDatas;
        }

        public void SaveWorkBook() {
            this.excelStream.Close();
            FileStream closeStream = new FileStream(this.filePath, FileMode.Create, FileAccess.Write);
            this.workBook.Write(closeStream);
            Console.WriteLine(this.workBook.GetSheetAt(0).LastRowNum);
            closeStream.Close();
        }

        public void addRow(List<String> values) {
            int rowNum = this.workSheet.LastRowNum;
            IRow row=this.workSheet.CreateRow(rowNum);
            for (int i = 0; i < values.Count; i++) {
                row.CreateCell(i).SetCellValue(values[i]);
            }
        }
        public void addColumn(List<String> values, String title) {
            int maxColumnNum = ((IEnumerable<IRow>)this.workSheet).Select(i => i.LastCellNum).Max();

        }
    }
}
