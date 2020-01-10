using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ConsoleApplication1.Orcl_ldbs3;
using System.Threading.Tasks;
using System.Data;
using System.Reflection;
using System.Dynamic;

namespace ConsoleApplication1
{
    class BaseData {
        public String data;
    }

    class WaterLevelData {
        public BaseData 站名;
        public BaseData ID号;
        public BaseData 经度;
        public BaseData 纬度;
        public BaseData 数据来源;
        public BaseData 区域归属;
        public BaseData 类型;
        public String 水位;
        public String 警戒水位;
        public String STNM
        {
            get { return STNM; }
            set {
                站名 = new BaseData() { data=value};
        } }
        public String STCD
        {
            get { return STCD; }
            set {
                ID号 = new BaseData() { data = value };
        } }
        public String LGTD
        {
            get { return LGTD; }
            set {
                经度 = new BaseData() { data = value };
        } }
        public String LTTD
        {
            get { return LTTD; }
            set {
                纬度 = new BaseData() { data = value };
        } }
        public String FRGRD
        {
            get { return FRGRD; }
            set {
                数据来源 = new BaseData() { data = value };
        } }
        public String BSNM
        {
            get { return BSNM; }
            set {
                区域归属 = new BaseData() { data = value };
        } }
        public String STTP
        {
            get { return STTP; }
            set {
                类型 = new BaseData() { data = value };
            } 
        }
        public String Z {
            get { return STTP; }
            set
            {
                水位 = value;
            }
        }
        public String WRZ {
            get { return STTP; }
            set
            {
                警戒水位 = value;
            }
        }
    }

    class Program
    {

        static public List<T> Tolist<T>(DataTable dt) where T : class, new()
        {
            Type t = typeof(T);
            PropertyInfo[] PropertyInfo = t.GetProperties();
            List<T> list = new List<T>();

            string typeName = string.Empty;
            foreach (DataRow item in dt.Rows)
            {
                T obj = new T();
                foreach (PropertyInfo s in PropertyInfo)
                {
                    typeName = s.Name;
                    //Console.WriteLine(typeName);
                    if (dt.Columns.Contains(typeName))
                    {
                        //Console.WriteLine("enter");
                        if (!s.CanWrite) continue;

                        object value = item[typeName];
                        if (value == DBNull.Value) continue;

                        if (s.PropertyType == typeof(string))
                        {
                            s.SetValue(obj, value.ToString(), null);
                            //Console.WriteLine("enter string"+obj+value.ToString());
                        }
                        else
                        {
                            s.SetValue(obj, value, null);
                        }
                    }
                }
                list.Add(obj);
            }
            return list;
        }

        static public List<dynamic> extractBaseData<T>(List<T> dataList) {
            var BaseDatas = dataList.Select(r =>
            {
                Type R = typeof(T);
                dynamic d = new DynamicModel();
                foreach (var field in R.GetFields())
                {
                    if (field.FieldType.Name == "BaseData")
                    {
                        d.PropertyName = field.Name;
                        dynamic value = field.GetValue(r);
                        d.Property = value.data;
                    }
                }
                return d;
            }).ToList();
            return BaseDatas;
        }

        static void Main(string[] args)
        {
            Orcl_Idbs3 orcl = new Orcl_Idbs3();
            orcl.GetDtTableXYCompleted+=(o,e)=>{
                DataTable dt=e.Result;
                List<WaterLevelData> WaterLevelDatas = Tolist<WaterLevelData>(dt);
                Console.WriteLine(WaterLevelDatas.Count);
                WaterLevelDatas = (from r in WaterLevelDatas
                                  group r by r.ID号 into n
                                  select n.ToArray()[0]).ToList();
                var BaseDatas = extractBaseData(WaterLevelDatas);
                OperateExcel oe = new OperateExcel(@"C:\Users\24513\Desktop\10.xlsx");
                oe.CreateOrOpenWorkSheet("test");
                oe.SerializableBaseDatas(BaseDatas);
                //oe.SaveWorkBook();
                var readBaseDatas = oe.DeSerialableBaseDatas(new List<String>() { "ID号","站名" });
                oe.addColumn(null, "");
                Console.WriteLine("read" + readBaseDatas.Count);
                oe.SaveWorkBook();
                //WaterLevelData[] AppendDatas=new WaterLevelData[WaterLevelDatas.Count];
                //WaterLevelDatas.CopyTo(AppendDatas, 0);
                //var AppendBaseDatas = AppendDatas.Select(r =>
                //{
                //    Type R = typeof(WaterLevelData);
                //    dynamic d = new DynamicModel();
                //    foreach (var field in R.GetFields())
                //    {
                //        if (field.FieldType.Name == "BaseData")
                //        {
                //            d.PropertyName = field.Name;
                //            dynamic value = field.GetValue(r);
                //            d.Property = value.data;
                //        }
                //    }
                //    return d;
                //}).ToList();
                //Console.WriteLine(BaseDatas[1].ID号);
                //BaseDatas.RemoveAt(1);
                //BaseDatas.RemoveAt(1);
                //Console.WriteLine(BaseDatas.Count + "," + AppendBaseDatas.Count);
                //var appendlist = from r in AppendBaseDatas
                //                 join a in BaseDatas on r.ID号 equals a.ID号 into match
                //                 where !match.Any()
                //                 select r;
                //BaseDatas = BaseDatas.Concat(appendlist).ToList();
                //var finalDatas = from r in BaseDatas
                //            join a in AppendDatas on r.ID号 equals a.ID号.data
                //            select new { r, a.水位 };
                //foreach (var data in finalDatas) {
                //    Console.WriteLine(data.水位);
                //}
            };
            orcl.GetDtTableXYAsync("1","2019-12-29","","","","");
            Console.ReadLine();
        }

        static void ExportExcel(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                return;
            }
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range range;
            long totalCount = dt.Rows.Count;
            long rowRead = 0;
            float percent = 0;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, i + 1];
                range.Interior.ColorIndex = 15;
                range.Font.Bold = true;
            }
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = dt.Rows[r][i].ToString();
                }
                rowRead++;
                percent = ((float)(100 * rowRead)) / totalCount;
            }
            xlApp.Visible = true;
        }
    }
}
