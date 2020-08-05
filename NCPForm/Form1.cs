using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using NPOI.SS.Util;

namespace NCP
{
    public partial class Form1 : Form
    {
        private static SqLiteHelper SqLiteHelper;
        private static string POS = ConfigurationManager.AppSettings["POS"];
        private static string HEX = ConfigurationManager.AppSettings["HEX"];
        private static string FAM = ConfigurationManager.AppSettings["FAM"];
        private static string ROX = ConfigurationManager.AppSettings["ROX"];
        Model model;
        
        public Form1()
        {
            InitializeComponent();
            SqLiteHelper = new SqLiteHelper("data source=NCPDB.db");

            // SqLiteHelper.InsertValues("CT_DB", new string[]{ "A03", "SYBR", "786/787", "Unkn","", "31.5518958512837", "31.5518958512837", "0", "60",DateTime.Now.ToString("yyyy-mm-dd hh:mm:ss") });
            //getExeclinfo();
        }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }

        public DataTable getExeclinfo()
        {
            DataTable dt = new DataTable();
            //打开文件，获取execl文件
            OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                if (fd.FileName.Contains(".xls") && fd.FileName.Contains(".xlsx"))
                {
                    var fileName = fd.FileName;
                    dt = ExcelToTable(fileName);
                }
                else
                {
                    MessageBox.Show("请选择execl文件");
                }
            }
            return dt;
        }
        private DataTable getExeclinfo(ref string datetime) 
        {
            DataTable dt = new DataTable();
            //打开文件，获取execl文件
            OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                if (fd.FileName.Contains(".xls") && fd.FileName.Contains(".xlsx"))
                {
                    var fileName = fd.FileName;
                    dt = ExcelToTable(fileName, ref datetime);
                }
                else
                {
                    MessageBox.Show("请选择execl文件");
                }
            }
            return dt;
        }

        /// <summary>
        /// Excel导入成DataTble(RUF)
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(obj.ToString());
                    }

                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
                //sheet2
                ISheet sheet2 = workbook.GetSheetAt(1);

               
                //数据  
                for (int i = sheet2.FirstRowNum + 1; i <= sheet2.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet2.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
                
                //sheet3
                ISheet sheet3 = workbook.GetSheetAt(2);
                //数据  
                for (int i = sheet3.FirstRowNum + 1; i <= sheet3.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet3.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }

            }
            return dt;
        }
        public static DataTable ExcelToTable(string file,ref string datetime)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);
                //读取检测时间
                ISheet sheet2 = workbook.GetSheetAt(1);
               
                datetime = sheet2.GetRow(sheet2.FirstRowNum+5).GetCell(1).ToString().Split(new string[]{" UTC"},StringSplitOptions.RemoveEmptyEntries)[0];
                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(obj.ToString());
                    }

                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                default:
                    return "=" + cell.CellFormula;
            }
        }


        /// <summary>
        /// 导入CT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnct_Click(object sender, EventArgs e)
        {
            try
            {
                var sql = "INSERT INTO CT_DB ";
                var datetime = string.Empty;
                var datatable = getExeclinfo(ref datetime);

                if (datatable.Rows.Count < 1)
                {
                    return;
                }
                if (!SqLiteHelper.FieldCountIsOk("CT_DB", datatable.Columns.Count + 2))
                {
                    MessageBox.Show("CT数据异常，请检查数据！");
                    return;
                }
                //求出当前excel每个通道的POS的值
                var hexPosRows = datatable.Select("Fluor='" + HEX + "'  and Content='" + POS + "'");
                decimal sumhexPosRows = 0;
                for (int i = 0; i < hexPosRows.Count(); i++)
                {
                    sumhexPosRows += decimal.Parse(hexPosRows[i]["Cq"].ToString());
                }
                decimal hexPos = sumhexPosRows / hexPosRows.Count();

                var famPosRows = datatable.Select("Fluor='" + FAM + "'  and Content='" + POS + "'");
                decimal sumfamPosRows = 0;
                for (int i = 0; i < famPosRows.Count(); i++)
                {
                    sumfamPosRows += decimal.Parse(famPosRows[i]["Cq"].ToString());
                }
                decimal famPos = sumfamPosRows / famPosRows.Count();

                var roxPosRows = datatable.Select("Fluor='" + ROX + "'  and Content='" + POS + "'");
                decimal sumroxPosRows = 0;
                for (int i = 0; i < roxPosRows.Count(); i++)
                {
                    sumroxPosRows += decimal.Parse(roxPosRows[i]["Cq"].ToString());
                }
                decimal roxPos = sumroxPosRows / roxPosRows.Count();

                //Cq Mean保存pos值

                for (int i = 0; i < datatable.Rows.Count; i++)
                {
                    //Sample最后一位是类型，前面是ID，Content还是存ID号（改动小点）
                    var ctType = string.Empty;
                    var cq = string.Empty;
                    if (datatable.Rows[i][4].ToString().ToUpper() == POS)
                    {
                        cq = datatable.Rows[i][4].ToString().ToUpper();
                    }
                    else
                    {
                        cq = datatable.Rows[i][5].ToString().Substring(0, datatable.Rows[i][5].ToString().Length - 1);
                        ctType = datatable.Rows[i][5].ToString().Substring(datatable.Rows[i][5].ToString().Length - 1);
                    }
                    if (i == datatable.Rows.Count - 1)
                    {

                        if (datatable.Rows[i][2].ToString().ToUpper() == HEX)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), hexPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }
                        else if (datatable.Rows[i][2].ToString().ToUpper() == FAM)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), famPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }
                        else
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), roxPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }

                        //sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);
                    }
                    else
                    {
                        //sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}'  UNION ALL ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);
                        if (datatable.Rows[i][2].ToString().ToUpper() == HEX)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}' UNION ALL  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datetime, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), hexPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }
                        else if (datatable.Rows[i][2].ToString().ToUpper() == FAM)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}' UNION ALL  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), famPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }
                        else
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}' ,'{17}' UNION ALL  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), roxPos, datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), datatable.Rows[i][11].ToString(), datatable.Rows[i][12].ToString(), datatable.Rows[i][13].ToString(), datatable.Rows[i][14].ToString(), datatable.Rows[i][15].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), ctType);

                        }
                    }
                }
                SqLiteHelper.ExecuteQuery(sql);
                MessageBox.Show("导入CT数据成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("导入CT数据错误，" + ex.Message);
            }
        }

        private void btnrfu_Click(object sender, EventArgs e)
        {
            try
            {
                var sql = "INSERT INTO RFU_DB ";
                var datatable = getExeclinfo();
                if (datatable.Rows.Count < 1)
                {
                    return;
                }
                if (!SqLiteHelper.FieldCountIsOk("RFU_DB", datatable.Columns.Count + 2))
                {
                    MessageBox.Show("RFU数据异常，请检查数据！");
                    return;
                }
                //求出当前excel每个通道的POS的值
                var hexPosRows = datatable.Select("Fluor='" + HEX + "'  and Content='" + POS + "'");
                decimal sumhexPosRows = 0;
                for (int i = 0; i < hexPosRows.Count(); i++)
                {
                    sumhexPosRows += decimal.Parse(hexPosRows[i]["End RFU"].ToString());
                }
                decimal hexPos = sumhexPosRows / hexPosRows.Count();

                var famPosRows = datatable.Select("Fluor='" + FAM + "'  and Content='" + POS + "'");
                decimal sumfamPosRows = 0;
                for (int i = 0; i < famPosRows.Count(); i++)
                {
                    sumfamPosRows += decimal.Parse(famPosRows[i]["End RFU"].ToString());
                }
                decimal famPos = sumfamPosRows / famPosRows.Count();

                var roxPosRows = datatable.Select("Fluor='" + ROX + "'  and Content='" + POS + "'");
                decimal sumroxPosRows = 0;
                for (int i = 0; i < roxPosRows.Count(); i++)
                {
                    sumroxPosRows += decimal.Parse(roxPosRows[i]["End RFU"].ToString());
                }
                decimal roxPos = sumroxPosRows / roxPosRows.Count();

                for (int i = 0; i < datatable.Rows.Count; i++)
                {
                    var rfuType = string.Empty;
                    var cq = string.Empty;
                    if (datatable.Rows[i][4].ToString().ToUpper() == POS)
                    {
                        cq = datatable.Rows[i][4].ToString().ToUpper();
                    }
                    else
                    {
                        cq = datatable.Rows[i][5].ToString().Substring(0, datatable.Rows[i][5].ToString().Length - 1);
                        rfuType = datatable.Rows[i][5].ToString().Substring(datatable.Rows[i][5].ToString().Length - 1);
                    }
                    if (i == datatable.Rows.Count - 1)
                    {
                        if (datatable.Rows[i][2].ToString().ToUpper() == HEX) 
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), hexPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                       else if (datatable.Rows[i][2].ToString().ToUpper() == FAM)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), famPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                        else 
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), roxPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                        //sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'  ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);
                    
                    }
                    else
                    {
                        if (datatable.Rows[i][2].ToString().ToUpper() == HEX)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}' UNION ALL   ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), hexPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                        else if (datatable.Rows[i][2].ToString().ToUpper() == FAM)
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}' UNION ALL   ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), famPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                        else
                        {
                            sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}' UNION ALL   ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), roxPos, cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);

                        }
                        // sql += string.Format("SELECT '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'   UNION ALL ", datatable.Rows[i][1].ToString(), datatable.Rows[i][2].ToString(), datatable.Rows[i][3].ToString(), cq, datatable.Rows[i][5].ToString(), datatable.Rows[i][6].ToString(), datatable.Rows[i][7].ToString(), datatable.Rows[i][8].ToString(), datatable.Rows[i][9].ToString(), datatable.Rows[i][10].ToString(), DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), rfuType);
                    }
                }
                SqLiteHelper.ExecuteQuery(sql);
                MessageBox.Show("导入RFU数据成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("导入RFU数据错误，" + ex.Message);
            }
        }
        /// <summary>
        /// 生成报表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btncrate_Click(object sender, EventArgs e)
        {
            CrateTable();

        }

        private void btnexecl_Click(object sender, EventArgs e)
        {
            var savePath = System.Environment.CurrentDirectory + @"\excel\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            if (model == null)
            {
                MessageBox.Show("请先生成报表，再导出报表！");
                return;
            }
            var wk = new XSSFWorkbook();
            //加粗样式,右对齐
            var style = wk.CreateCellStyle();
            IFont f = wk.CreateFont();
            f.Boldweight = (short)FontBoldWeight.Bold;
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            style.VerticalAlignment = VerticalAlignment.Center;
            f.FontHeightInPoints = 10;
            style.SetFont(f);
            //居中,边框
            ICellStyle style2 = wk.CreateCellStyle();
            style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style2.VerticalAlignment = VerticalAlignment.Center;
            style2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            IFont f2 = wk.CreateFont();
            f2.FontHeightInPoints = 10;
            style2.SetFont(f2);

            //居中,边框,背景
            ICellStyle style3 = wk.CreateCellStyle();
            style3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style3.VerticalAlignment = VerticalAlignment.Center;
            style3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style3.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style3.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style3.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            style3.FillPattern = FillPattern.SolidForeground;
            IFont f3 = wk.CreateFont();
            f3.FontHeightInPoints = 10;
            style3.SetFont(f3);

            //居中,边框，红色
            ICellStyle style4 = wk.CreateCellStyle();
            style4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style4.VerticalAlignment = VerticalAlignment.Center;
            style4.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style4.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style4.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style4.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            //style4.Rotation = (short)0xff;
            IFont f4 = wk.CreateFont();
            f4.FontHeightInPoints = 10;
            f4.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
            style4.SetFont(f4);

            //居中,边框，文字竖排
            ICellStyle style5 = wk.CreateCellStyle();
            style5.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style5.VerticalAlignment = VerticalAlignment.Center;
            style5.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style5.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style5.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style5.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style5.Rotation = (short)0xff;
            IFont f5 = wk.CreateFont();
            f5.FontHeightInPoints = 10;
            style5.SetFont(f5);

            //声明工作表
            var sheet = wk.CreateSheet();
            //sheet.DisplayGridlines = false;//设置默认为无边框
            //sheet.DefaultRowHeight = 30 * 50;
            //(Optional) set the width of the columns
            sheet.SetColumnWidth(0, 10 * 256);

            sheet.SetColumnWidth(1, 13 * 256);
            sheet.SetColumnWidth(2, 13 * 256);
            sheet.SetColumnWidth(3, 13 * 256);
            sheet.SetColumnWidth(4, 13 * 256);
            sheet.SetColumnWidth(5, 13 * 256);
            sheet.SetColumnWidth(6, 13 * 256);
            sheet.SetColumnWidth(7, 13 * 256);


            //创建行(默认从0行开始)
            var headerRow = sheet.CreateRow(0);
            ////创建单元格(默认从0行开始)
            //var c = headerRow.CreateCell(0);
            var row1 = sheet.CreateRow(1);
            row1.Height = 25 * 20;
            row1.CreateCell(1).SetCellValue("姓名:");
            sheet.GetRow(1).GetCell(1).CellStyle = style;
            row1.CreateCell(3).SetCellValue("ID号:");
            sheet.GetRow(1).GetCell(3).CellStyle = style;
            row1.CreateCell(4).SetCellValue(model.batchNo);
            row1.CreateCell(5).SetCellValue("检测时间:");
            sheet.GetRow(1).GetCell(5).CellStyle = style;
            row1.CreateCell(6).SetCellValue(model.Datetime);
            //row2

            //合并单元格
            /**
              第一个参数：从第几行开始合并
              第二个参数：到第几行结束合并
              第三个参数：从第几列开始合并
              第四个参数：到第几列结束合并
          **/
            CellRangeAddress region = new CellRangeAddress(2, 2, 0, 2);
            sheet.AddMergedRegion(region);
            for (int i = 2; i < 17; i++)
            {
                if (i % 2 == 0 || i == 3)
                {
                    var row = sheet.CreateRow(i);
                    row.Height = 25 * 20;
                    if (i == 16)//设置字体红色
                    {
                        for (int j = 0; j <= 7; j++)
                        {

                            var cel = row.CreateCell(j);
                            cel.SetCellValue("");
                            row.GetCell(j).CellStyle = style4;

                        }
                    }
                    else
                    {
                        for (int j = 0; j <= 7; j++)
                        {

                            var cel = row.CreateCell(j);
                            cel.SetCellValue("");
                            row.GetCell(j).CellStyle = style2;

                        }
                    }


                }
                else
                {
                    var row = sheet.CreateRow(i);
                    row.Height = 25 * 20;
                    for (int j = 0; j <= 7; j++)
                    {

                        var cel = row.CreateCell(j);
                        cel.SetCellValue("");
                        row.GetCell(j).CellStyle = style3;

                    }
                }
            }


            sheet.GetRow(2).GetCell(0).SetCellValue("样本类型");
            sheet.GetRow(2).GetCell(3).SetCellValue("阳性质控");
            sheet.GetRow(2).GetCell(4).SetCellValue("咽拭子/A");
            sheet.GetRow(2).GetCell(5).SetCellValue("肛拭子/B");
            sheet.GetRow(2).GetCell(6).SetCellValue("痰液/C");
            sheet.GetRow(2).GetCell(7).SetCellValue("肺灌洗液/D");
            //row3
            CellRangeAddress region2 = new CellRangeAddress(3, 3, 0, 2);
            sheet.AddMergedRegion(region2);
            sheet.GetRow(3).GetCell(0).SetCellValue("平行送检样本");
            sheet.GetRow(3).GetCell(3).SetCellValue("");
            sheet.GetRow(3).GetCell(4).SetCellValue(model.lblIsCheck1);
            sheet.GetRow(3).GetCell(5).SetCellValue(model.lblIsCheck2);
            sheet.GetRow(3).GetCell(6).SetCellValue(model.lblIsCheck3);
            sheet.GetRow(3).GetCell(7).SetCellValue(model.lblIsCheck4);
            //row4
            sheet.GetRow(4).GetCell(0).SetCellValue("新冠病毒核酸检测结果");
            sheet.GetRow(4).GetCell(0).CellStyle = style5;
            sheet.GetRow(4).GetCell(1).SetCellValue("内标");
            sheet.GetRow(4).GetCell(2).SetCellValue("CT值");
            sheet.GetRow(4).GetCell(3).SetCellValue(model.lblsumctA.ToString());
            sheet.GetRow(4).GetCell(4).SetCellValue(model.lblct1.ToString());
            sheet.GetRow(4).GetCell(5).SetCellValue(model.lblct2.ToString());
            sheet.GetRow(4).GetCell(6).SetCellValue(model.lblct3.ToString());
            sheet.GetRow(4).GetCell(7).SetCellValue(model.lblct4.ToString());

            //row5
            // sheet.GetRow(5).GetCell(0).SetCellValue("人类内标看家基因");
            sheet.GetRow(5).GetCell(2).SetCellValue("△CT值");
            sheet.GetRow(5).GetCell(3).SetCellValue("-");
            sheet.GetRow(5).GetCell(4).SetCellValue(model.lblctA1.ToString());
            sheet.GetRow(5).GetCell(5).SetCellValue(model.lblctA2.ToString());
            sheet.GetRow(5).GetCell(6).SetCellValue(model.lblctA3.ToString());
            sheet.GetRow(5).GetCell(7).SetCellValue(model.lblctA4.ToString());

            //row6
            //sheet.GetRow(6).GetCell(0).SetCellValue("人类内标看家基因");
            sheet.GetRow(6).GetCell(2).SetCellValue("RFU值");
            sheet.GetRow(6).GetCell(3).SetCellValue(model.lblsumRFCA.ToString());
            sheet.GetRow(6).GetCell(4).SetCellValue(model.lblrfu1.ToString());
            sheet.GetRow(6).GetCell(5).SetCellValue(model.lblrfu2.ToString());
            sheet.GetRow(6).GetCell(6).SetCellValue(model.lblrfu3.ToString());
            sheet.GetRow(6).GetCell(7).SetCellValue(model.lblrfu4.ToString());

            //row7
            sheet.GetRow(7).GetCell(2).SetCellValue("△RFU值");
            sheet.GetRow(7).GetCell(3).SetCellValue("-");
            sheet.GetRow(7).GetCell(4).SetCellValue(model.lblrfuA1.ToString());
            sheet.GetRow(7).GetCell(5).SetCellValue(model.lblrfuA2.ToString());
            sheet.GetRow(7).GetCell(6).SetCellValue(model.lblrfuA3.ToString());
            sheet.GetRow(7).GetCell(7).SetCellValue(model.lblrfuA4.ToString());
            //合并
            CellRangeAddress region3 = new CellRangeAddress(4, 7, 1, 1);
            sheet.AddMergedRegion(region3);

            //row8
            //sheet.GetRow(8).GetCell(0).SetCellValue("新型冠状病毒");
            CellRangeAddress region6 = new CellRangeAddress(4, 15, 0, 0);
            sheet.AddMergedRegion(region6);
            sheet.GetRow(8).GetCell(1).SetCellValue("ORF1ab片段");
            sheet.GetRow(8).GetCell(2).SetCellValue("CT值");
            sheet.GetRow(8).GetCell(3).SetCellValue(model.lblsumctB.ToString());
            sheet.GetRow(8).GetCell(4).SetCellValue(model.lblORCT1.ToString());
            sheet.GetRow(8).GetCell(5).SetCellValue(model.lblORCT2.ToString());
            sheet.GetRow(8).GetCell(6).SetCellValue(model.lblORCT3.ToString());
            sheet.GetRow(8).GetCell(7).SetCellValue(model.lblORCT4.ToString());

            //row9
            sheet.GetRow(9).GetCell(2).SetCellValue("△CT值");
            sheet.GetRow(9).GetCell(3).SetCellValue("-");
            sheet.GetRow(9).GetCell(4).SetCellValue(model.lblctORA1.ToString());
            sheet.GetRow(9).GetCell(5).SetCellValue(model.lblctORA2.ToString());
            sheet.GetRow(9).GetCell(6).SetCellValue(model.lblctORA3.ToString());
            sheet.GetRow(9).GetCell(7).SetCellValue(model.lblctORA4.ToString());

            //row10
            sheet.GetRow(10).GetCell(2).SetCellValue("RFU值");
            sheet.GetRow(10).GetCell(3).SetCellValue(model.lblsumRFCB.ToString());
            sheet.GetRow(10).GetCell(4).SetCellValue(model.lblORrfu1.ToString());
            sheet.GetRow(10).GetCell(5).SetCellValue(model.lblORrfu2.ToString());
            sheet.GetRow(10).GetCell(6).SetCellValue(model.lblORrfu3.ToString());
            sheet.GetRow(10).GetCell(7).SetCellValue(model.lblORrfu4.ToString());

            //row11
            sheet.GetRow(11).GetCell(2).SetCellValue("△RFU值");
            sheet.GetRow(11).GetCell(3).SetCellValue("-");
            sheet.GetRow(11).GetCell(4).SetCellValue(model.lblORrfuA1.ToString());
            sheet.GetRow(11).GetCell(5).SetCellValue(model.lblORrfuA2.ToString());
            sheet.GetRow(11).GetCell(6).SetCellValue(model.lblORrfuA3.ToString());
            sheet.GetRow(11).GetCell(7).SetCellValue(model.lblORrfuA4.ToString());
            //合并
            CellRangeAddress region4 = new CellRangeAddress(8, 11, 1, 1);
            sheet.AddMergedRegion(region4);

            //row12
            sheet.GetRow(12).GetCell(1).SetCellValue("N片段");
            sheet.GetRow(12).GetCell(2).SetCellValue("CT值");
            sheet.GetRow(12).GetCell(3).SetCellValue(model.lblsumctC.ToString());
            sheet.GetRow(12).GetCell(4).SetCellValue(model.lblcq1.ToString());
            sheet.GetRow(12).GetCell(5).SetCellValue(model.lblcq2.ToString());
            sheet.GetRow(12).GetCell(6).SetCellValue(model.lblcq3.ToString());
            sheet.GetRow(12).GetCell(7).SetCellValue(model.lblcq4.ToString());
            //row13

            sheet.GetRow(13).GetCell(2).SetCellValue("△CT值");
            sheet.GetRow(13).GetCell(3).SetCellValue("-");
            sheet.GetRow(13).GetCell(4).SetCellValue(model.lblctNA1.ToString());
            sheet.GetRow(13).GetCell(5).SetCellValue(model.lblctNA2.ToString());
            sheet.GetRow(13).GetCell(6).SetCellValue(model.lblctNA3.ToString());
            sheet.GetRow(13).GetCell(7).SetCellValue(model.lblctNA4.ToString());

            //row14

            sheet.GetRow(14).GetCell(2).SetCellValue("RFU值");
            sheet.GetRow(14).GetCell(3).SetCellValue(model.lblsumRFCC.ToString());
            sheet.GetRow(14).GetCell(4).SetCellValue(model.lblNrfu1.ToString());
            sheet.GetRow(14).GetCell(5).SetCellValue(model.lblNrfu2.ToString());
            sheet.GetRow(14).GetCell(6).SetCellValue(model.lblNrfu3.ToString());
            sheet.GetRow(14).GetCell(7).SetCellValue(model.lblNrfu4.ToString());

            //row15
            sheet.GetRow(15).GetCell(2).SetCellValue("△RFU值");
            sheet.GetRow(15).GetCell(3).SetCellValue("-");
            sheet.GetRow(15).GetCell(4).SetCellValue(model.lblNrfuA1.ToString());
            sheet.GetRow(15).GetCell(5).SetCellValue(model.lblNrfuA2.ToString());
            sheet.GetRow(15).GetCell(6).SetCellValue(model.lblNrfuA3.ToString());
            sheet.GetRow(15).GetCell(7).SetCellValue(model.lblNrfuA4.ToString());
            //合并


            CellRangeAddress region5 = new CellRangeAddress(12, 15, 1, 1);
            sheet.AddMergedRegion(region5);


            ////row16
            //sheet.GetRow(16).GetCell(0).SetCellValue("阳性质控");
            //sheet.GetRow(16).GetCell(2).SetCellValue("CT值");
            //sheet.GetRow(16).GetCell(3).SetCellValue(model.lblsumct.ToString());
            //CellRangeAddress region7 = new CellRangeAddress(16, 17, 0, 1);
            //sheet.AddMergedRegion(region7);

            //CellRangeAddress region8 = new CellRangeAddress(16, 16, 3, 6);
            //sheet.AddMergedRegion(region8);

            ////row17
            //sheet.GetRow(17).GetCell(2).SetCellValue("RFU值");
            //sheet.GetRow(17).GetCell(3).SetCellValue(model.lblsumRFC.ToString());
            //CellRangeAddress region9 = new CellRangeAddress(17, 17, 3, 6);
            //sheet.AddMergedRegion(region9);

            //row18
            sheet.GetRow(16).GetCell(0).SetCellValue("结论");
            sheet.GetRow(16).GetCell(4).SetCellValue(model.Result1);
            sheet.GetRow(16).GetCell(5).SetCellValue(model.Result2);
            sheet.GetRow(16).GetCell(6).SetCellValue(model.Result3);
            sheet.GetRow(16).GetCell(7).SetCellValue(model.Result4);
            CellRangeAddress region10 = new CellRangeAddress(16, 16, 0, 3);
            sheet.AddMergedRegion(region10);
            using (FileStream fs = new FileStream(savePath, FileMode.Create))
            {
                wk.Write(fs);
            }

            MessageBox.Show("导出成功");

        }

        private void txtID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CrateTable();
            }
        }

        private void CrateTable()
        {
            if (string.IsNullOrWhiteSpace(txtID.Text))
            {
                MessageBox.Show("请输入样本号！");
                return;
            }
            var ctSql = "SELECT Fluor, Cq, Date,CT_Type,Cq Mean,Target from CT_DB WHERE  Content =  '" + txtID.Text + "' ORDER BY DateTime DESC";
            var rfuSql = "SELECT Fluor, RFU, Date,CT_Type,Target from RFU_DB WHERE  Content =  '" + txtID.Text + "' ORDER BY DateTime DESC";
            var ctTable = SqLiteHelper.Query(ctSql);
            var rfuTable = SqLiteHelper.Query(rfuSql);
            if (ctTable.Rows.Count < 1)
            {
                MessageBox.Show("没有找到CT相关数据，请上传！");
                return;
            }
            if (rfuTable.Rows.Count < 1)
            {
                MessageBox.Show("没有找到RFU相关数据，请上传！");
                return;
            }
            model = new Model();
            #region CT
            model.batchNo = txtID.Text;
            var HEXs = ctTable.Select("Fluor='" + HEX + "'  and Date='" + ctTable.Rows[0]["Date"] + "'");
            var FAMs = ctTable.Select("Fluor='" + FAM + "'  and Date='" + ctTable.Rows[0]["Date"] + "'");
            var ROXs = ctTable.Select("Fluor='" + ROX + "'  and Date='" + ctTable.Rows[0]["Date"] + "'");
            model.lblsumctA = decimal.Round(Convert.ToDecimal(HEXs[0]["Mean"]),2);
            model.lblsumctB = decimal.Round(Convert.ToDecimal(FAMs[0]["Mean"]), 2);
            model.lblsumctC = decimal.Round(Convert.ToDecimal(ROXs[0]["Mean"]), 2);
            model.Datetime = Convert.ToDateTime(HEXs[0]["Target"]).ToString("yyyy-MM-dd HH:mm");
            //ctpos
            //var sumctsql = "SELECT Cq from  CT_DB WHERE Content='" + POS + "' and Date='" + ctTable.Rows[0]["Date"] + "'";
            //var sunctTable = SqLiteHelper.Query(sumctsql);
            //if (sunctTable.Rows.Count < 1)
            //{
            //    MessageBox.Show("没有找到CT的POS相关数据，请确认！");
            //    model = null;
            //    return;
            //}
            //var ctpos = decimal.Round(decimal.Parse(sunctTable.Compute("sum(Cq)", "").ToString()) / sunctTable.Rows.Count, 2);
            //model.lblsumct = ctpos;


            //初始化数据
            model.lblIsCheck1 = "×";
            model.lblIsCheck2 = "×";
            model.lblIsCheck3 = "×";
            model.lblIsCheck4 = "×";
            //HEXs
            for (int i = HEXs.Count() - 1; i >= 0; i--)
            {
                switch (HEXs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblct1 = decimal.Round(decimal.Parse(HEXs[i]["Cq"].ToString()), 2);
                        model.lblIsCheck1 = "√";
                        break;
                    case "B":
                        model.lblct2 = decimal.Round(decimal.Parse(HEXs[i]["Cq"].ToString()), 2);
                        model.lblIsCheck2 = "√";
                        break;
                    case "C":
                        model.lblct3 = decimal.Round(decimal.Parse(HEXs[i]["Cq"].ToString()), 2);
                        model.lblIsCheck3 = "√";
                        break;
                    case "D":
                        model.lblct4 = decimal.Round(decimal.Parse(HEXs[i]["Cq"].ToString()), 2);
                        model.lblIsCheck4 = "√";
                        break;
                    default:
                        break;
                }
            }
            //model.lblct1 = decimal.Round(decimal.Parse(HEXs[0]["Cq"].ToString()), 2);
            //model.lblct2 = decimal.Round(decimal.Parse(HEXs[1]["Cq"].ToString()), 2);
            //model.lblct3 = decimal.Round(decimal.Parse(HEXs[2]["Cq"].ToString()), 2);
            //model.lblct4 = decimal.Round(decimal.Parse(HEXs[3]["Cq"].ToString()), 2);

            //FAM
            for (int i = FAMs.Count() - 1; i >= 0; i--)
            {
                switch (FAMs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblORCT1 = decimal.Round(decimal.Parse(FAMs[i]["Cq"].ToString()), 2);
                        break;
                    case "B":
                        model.lblORCT2 = decimal.Round(decimal.Parse(FAMs[i]["Cq"].ToString()), 2);
                        break;
                    case "C":
                        model.lblORCT3 = decimal.Round(decimal.Parse(FAMs[i]["Cq"].ToString()), 2);
                        break;
                    case "D":
                        model.lblORCT4 = decimal.Round(decimal.Parse(FAMs[i]["Cq"].ToString()), 2);
                        break;
                    default:
                        break;
                }
            }
            //model.lblORCT1 = decimal.Round(decimal.Parse(FAMs[0]["Cq"].ToString()), 2);
            //model.lblORCT2 = decimal.Round(decimal.Parse(FAMs[1]["Cq"].ToString()), 2);
            //model.lblORCT3 = decimal.Round(decimal.Parse(FAMs[2]["Cq"].ToString()), 2);
            //model.lblORCT4 = decimal.Round(decimal.Parse(FAMs[3]["Cq"].ToString()), 2);

            //ROX
            for (int i = ROXs.Count() - 1; i >= 0; i--)
            {
                switch (ROXs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblcq1 = decimal.Round(decimal.Parse(ROXs[i]["Cq"].ToString()), 2);
                        break;
                    case "B":
                        model.lblcq2 = decimal.Round(decimal.Parse(ROXs[i]["Cq"].ToString()), 2);
                        break;
                    case "C":
                        model.lblcq3 = decimal.Round(decimal.Parse(ROXs[i]["Cq"].ToString()), 2);
                        break;
                    case "D":
                        model.lblcq4 = decimal.Round(decimal.Parse(ROXs[i]["Cq"].ToString()), 2);
                        break;
                    default:
                        break;
                }
            }

            //model.lblcq1 = decimal.Round(decimal.Parse(ROXs[0]["Cq"].ToString()), 2);
            //model.lblcq2 = decimal.Round(decimal.Parse(ROXs[1]["Cq"].ToString()), 2);
            //model.lblcq3 = decimal.Round(decimal.Parse(ROXs[2]["Cq"].ToString()), 2);
            //model.lblcq4 = decimal.Round(decimal.Parse(ROXs[3]["Cq"].ToString()), 2);
            #endregion

            #region RFU
            var rfuHEXs = rfuTable.Select("Fluor='" + HEX + "' and Date='" + rfuTable.Rows[0]["Date"] + "'");
            var rfuFAMs = rfuTable.Select("Fluor='" + FAM + "' and Date='" + rfuTable.Rows[0]["Date"] + "'");
            var rfuROXs = rfuTable.Select("Fluor='" + ROX + "' and Date='" + rfuTable.Rows[0]["Date"] + "'");
            //ctpos
            //var sumrfusql = "SELECT RFU from  RFU_DB WHERE Content='" + POS + "' and Date='" + rfuTable.Rows[0]["Date"] + "'";
            //var sunrfuTable = SqLiteHelper.Query(sumrfusql);
            //if (sunrfuTable.Rows.Count < 1)
            //{
            //    MessageBox.Show("没有找到RFU的POS相关数据，请确认！");
            //    model = null;
            //    return;
            //}
            //var rfupos = decimal.Round(decimal.Parse(sunrfuTable.Compute("sum(RFU)", "").ToString()) / sunrfuTable.Rows.Count, 2);
            //model.lblsumRFC = rfupos;
            model.lblsumRFCA = decimal.Round(Convert.ToDecimal(rfuHEXs[0]["Target"]),2);
            model.lblsumRFCB = decimal.Round(Convert.ToDecimal(rfuFAMs[0]["Target"]), 2);
            model.lblsumRFCC = decimal.Round(Convert.ToDecimal(rfuROXs[0]["Target"]), 2);
            //HEXs
            for (int i = rfuHEXs.Count() - 1; i >= 0; i--)
            {
                switch (rfuHEXs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblrfu1 = decimal.Round(decimal.Parse(rfuHEXs[i]["RFU"].ToString()), 2);
                        break;
                    case "B":
                        model.lblrfu2 = decimal.Round(decimal.Parse(rfuHEXs[i]["RFU"].ToString()), 2);
                        break;
                    case "C":
                        model.lblrfu3 = decimal.Round(decimal.Parse(rfuHEXs[i]["RFU"].ToString()), 2);

                        break;
                    case "D":
                        model.lblrfu4 = decimal.Round(decimal.Parse(rfuHEXs[i]["RFU"].ToString()), 2);
                        break;
                    default:
                        break;
                }
            }
            //model.lblrfu1 = decimal.Round(decimal.Parse(rfuHEXs[0]["RFU"].ToString()), 2);
            //model.lblrfu2 = decimal.Round(decimal.Parse(rfuHEXs[1]["RFU"].ToString()), 2);
            //model.lblrfu3 = decimal.Round(decimal.Parse(rfuHEXs[2]["RFU"].ToString()), 2);
            //model.lblrfu4 = decimal.Round(decimal.Parse(rfuHEXs[3]["RFU"].ToString()), 2);
            //FAM

            for (int i = rfuFAMs.Count() - 1; i >= 0; i--)
            {
                switch (rfuFAMs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblORrfu1 = decimal.Round(decimal.Parse(rfuFAMs[i]["RFU"].ToString()), 2);
                        break;
                    case "B":
                        model.lblORrfu2 = decimal.Round(decimal.Parse(rfuFAMs[i]["RFU"].ToString()), 2);
                        break;
                    case "C":
                        model.lblORrfu3 = decimal.Round(decimal.Parse(rfuFAMs[i]["RFU"].ToString()), 2);

                        break;
                    case "D":
                        model.lblORrfu4 = decimal.Round(decimal.Parse(rfuFAMs[i]["RFU"].ToString()), 2);
                        break;
                    default:
                        break;
                }
            }
            //model.lblORrfu1 = decimal.Round(decimal.Parse(rfuFAMs[0]["RFU"].ToString()), 2);
            //model.lblORrfu2 = decimal.Round(decimal.Parse(rfuFAMs[1]["RFU"].ToString()), 2);
            //model.lblORrfu3 = decimal.Round(decimal.Parse(rfuFAMs[2]["RFU"].ToString()), 2);
            //model.lblORrfu4 = decimal.Round(decimal.Parse(rfuFAMs[3]["RFU"].ToString()), 2);
            //ROX


            for (int i = rfuROXs.Count() - 1; i >= 0; i--)
            {
                switch (rfuROXs[i]["CT_Type"].ToString().ToUpper())
                {
                    case "A":
                        model.lblNrfu1 = decimal.Round(decimal.Parse(rfuROXs[i]["RFU"].ToString()), 2);
                        break;
                    case "B":
                        model.lblNrfu2 = decimal.Round(decimal.Parse(rfuROXs[i]["RFU"].ToString()), 2);
                        break;
                    case "C":
                        model.lblNrfu3 = decimal.Round(decimal.Parse(rfuROXs[i]["RFU"].ToString()), 2);

                        break;
                    case "D":
                        model.lblNrfu4 = decimal.Round(decimal.Parse(rfuROXs[i]["RFU"].ToString()), 2);
                        break;
                    default:
                        break;
                }
            }

            //  model.lblNrfu1 = decimal.Round(decimal.Parse(rfuROXs[0]["RFU"].ToString()), 2);
            //model.lblNrfu2 = decimal.Round(decimal.Parse(rfuROXs[1]["RFU"].ToString()), 2);
            //model.lblNrfu3 = decimal.Round(decimal.Parse(rfuROXs[2]["RFU"].ToString()), 2);
            //model.lblNrfu4 = decimal.Round(decimal.Parse(rfuROXs[3]["RFU"].ToString()), 2);
            #endregion
            //更新界面
            lblct1.Text = model.lblct1.ToString();
            lblct2.Text = model.lblct2.ToString();
            lblct3.Text = model.lblct3.ToString();
            lblct4.Text = model.lblct4.ToString();

            lblctA1.Text = model.lblctA1.ToString();
            lblctA2.Text = model.lblctA2.ToString();
            lblctA3.Text = model.lblctA3.ToString();
            lblctA4.Text = model.lblctA4.ToString();

            lblrfu1.Text = model.lblrfu1.ToString();
            lblrfu2.Text = model.lblrfu2.ToString();
            lblrfu3.Text = model.lblrfu3.ToString();
            lblrfu4.Text = model.lblrfu4.ToString();

            lblrfuA1.Text = model.lblrfuA1.ToString();
            lblrfuA2.Text = model.lblrfuA2.ToString();
            lblrfuA3.Text = model.lblrfuA3.ToString();
            lblrfuA4.Text = model.lblrfuA4.ToString();

            lblORCT1.Text = model.lblORCT1.ToString();
            lblORCT2.Text = model.lblORCT2.ToString();
            lblORCT3.Text = model.lblORCT3.ToString();
            lblORCT4.Text = model.lblORCT4.ToString();

            lblctORA1.Text = model.lblctORA1.ToString();
            lblctORA2.Text = model.lblctORA2.ToString();
            lblctORA3.Text = model.lblctORA3.ToString();
            lblctORA4.Text = model.lblctORA4.ToString();

            lblORrfu1.Text = model.lblORrfu1.ToString();
            lblORrfu2.Text = model.lblORrfu2.ToString();
            lblORrfu3.Text = model.lblORrfu3.ToString();
            lblORrfu4.Text = model.lblORrfu4.ToString();

            lblORrfuA1.Text = model.lblORrfuA1.ToString();
            lblORrfuA2.Text = model.lblORrfuA2.ToString();
            lblORrfuA3.Text = model.lblORrfuA3.ToString();
            lblORrfuA4.Text = model.lblORrfuA4.ToString();

            lblcq1.Text = model.lblcq1.ToString();
            lblcq2.Text = model.lblcq2.ToString();
            lblcq3.Text = model.lblcq3.ToString();
            lblcq4.Text = model.lblcq4.ToString();

            lblctNA1.Text = model.lblctNA1.ToString();
            lblctNA2.Text = model.lblctNA2.ToString();
            lblctNA3.Text = model.lblctNA3.ToString();
            lblctNA4.Text = model.lblctNA4.ToString();

            lblNrfu1.Text = model.lblNrfu1.ToString();
            lblNrfu2.Text = model.lblNrfu2.ToString();
            lblNrfu3.Text = model.lblNrfu3.ToString();
            lblNrfu4.Text = model.lblNrfu4.ToString();

            lblNrfuA1.Text = model.lblNrfuA1.ToString();
            lblNrfuA2.Text = model.lblNrfuA2.ToString();
            lblNrfuA3.Text = model.lblNrfuA3.ToString();
            lblNrfuA4.Text = model.lblNrfuA4.ToString();

            lblsumctA.Text = model.lblsumctA.ToString();
            lblsumctB.Text = model.lblsumctB.ToString();
            lblsumctC.Text = model.lblsumctC.ToString();

            lblsumRFCA.Text = model.lblsumRFCA.ToString();
            lblsumRFCB.Text = model.lblsumRFCB.ToString();
            lblsumRFCC.Text = model.lblsumRFCC.ToString();
            lblIsCheck1.Text = model.lblIsCheck1.ToString();
            lblIsCheck2.Text = model.lblIsCheck2.ToString();
            lblIsCheck3.Text = model.lblIsCheck3.ToString();
            lblIsCheck4.Text = model.lblIsCheck4.ToString();
            //阳性(阳性判断标准：FAM或ROX通道0<CT≤40且RFU>500)
            if ((0 < model.lblORCT1 && model.lblORCT1 <= 40 && model.lblORrfu1 > 500) || (0 < model.lblcq1 && model.lblcq1 <= 40 && model.lblNrfu1 > 500))
            {
                model.Result1 = "阳性";
            }
            else if (model.lblct1 > 0 && model.lblct1 <= 40 && model.lblORCT1 <= 0 && model.lblcq1 <= 0)
            {
                model.Result1 = "阴性";
            }
            else
            {
                model.Result1 = "待复查";
            }
            if ((0 < model.lblORCT2 && model.lblORCT2 <= 40 && model.lblORrfu2 > 500) || (0 < model.lblcq2 && model.lblcq2 <= 40 && model.lblNrfu2 > 500))
            {
                model.Result2 = "阳性";
            }
            else if (model.lblct2 > 0 && model.lblct2 <= 40 && model.lblORCT2 <= 0 && model.lblcq2 <= 0)
            {
                model.Result2 = "阴性";
            }
            else
            {
                model.Result2 = "待复查";
            }

            if ((0 < model.lblORCT3 && model.lblORCT3 <= 40 && model.lblORrfu3 > 500) || (0 < model.lblcq3 && model.lblcq3 <= 40 && model.lblNrfu3 > 500))
            {
                model.Result3 = "阳性";
            }
            else if (model.lblct3 > 0 && model.lblct3 <= 40 && model.lblORCT3 <= 0 && model.lblcq3 <= 0)
            {
                model.Result3 = "阴性";
            }
            else
            {
                model.Result3 = "待复查";
            }

            if ((0 < model.lblORCT4 && model.lblORCT4 <= 40 && model.lblORrfu4 > 500) || (0 < model.lblcq4 && model.lblcq4 <= 40 && model.lblNrfu4 > 500))
            {
                model.Result4 = "阳性";
            }

            else if (model.lblct4 > 0 && model.lblct4 <= 40 && model.lblORCT4 <= 0 && model.lblcq4 <= 0)
            {
                model.Result4 = "阴性";
            }
            else
            {
                model.Result4 = "待复查";
            }
            //阴性(阴性判断标准：在HEX通道0<CT≤40条件下，同时满足FAM通道CT≤0且ROX通道CT≤0)




            //待复查：1.HEX通道CT≤0或CT＞40；2.FAM或ROX通道0<CT≤40且RFU≤500；3.ROX通道荧光信号有明显增幅，但CT＞40.


            txtResult1.Text = model.Result1;
            txtResult2.Text = model.Result2;
            txtResult3.Text = model.Result3;
            txtResult4.Text = model.Result4;

            tableLayoutPanel1.Visible = true;

        }


        public class Model
        {

            public string batchNo { get; set; }
            public string lblIsCheck1 { get; set; }
            public string lblIsCheck2 { get; set; }
            public string lblIsCheck3 { get; set; }
            public string lblIsCheck4 { get; set; }
            //HEXs_CT
            public decimal? lblct1 { get; set; }
            public decimal? lblct2 { get; set; }
            public decimal? lblct3 { get; set; }
            public decimal? lblct4 { get; set; }
            //ΔHEXs_CT
            public decimal? lblctA1 { get { return lblct1 - lblsumctA; } }
            public decimal? lblctA2 { get { return lblct2 - lblsumctA; } }
            public decimal? lblctA3 { get { return lblct3 - lblsumctA; } }
            public decimal? lblctA4 { get { return lblct4 - lblsumctA; } }
            //HEXs_RFU
            public decimal? lblrfu1 { get; set; }
            public decimal? lblrfu2 { get; set; }
            public decimal? lblrfu3 { get; set; }
            public decimal? lblrfu4 { get; set; }
            //ΔHEXs_RFU
            public decimal? lblrfuA1 { get { return lblrfu1 - lblsumRFCA; } }
            public decimal? lblrfuA2 { get { return lblrfu2 - lblsumRFCA; } }
            public decimal? lblrfuA3 { get { return lblrfu3 - lblsumRFCA; } }
            public decimal? lblrfuA4 { get { return lblrfu4 - lblsumRFCA; } }

            //FAM_CT
            public decimal? lblORCT1 { get; set; }
            public decimal? lblORCT2 { get; set; }
            public decimal? lblORCT3 { get; set; }
            public decimal? lblORCT4 { get; set; }
            //ΔFAM_CT
            public decimal? lblctORA1 { get { return lblORCT1 - lblsumctB; } }
            public decimal? lblctORA2 { get { return lblORCT2 - lblsumctB; } }
            public decimal? lblctORA3 { get { return lblORCT3 - lblsumctB; } }
            public decimal? lblctORA4 { get { return lblORCT4 - lblsumctB; } }
            //FAM_RFU
            public decimal? lblORrfu1 { get; set; }
            public decimal? lblORrfu2 { get; set; }
            public decimal? lblORrfu3 { get; set; }
            public decimal? lblORrfu4 { get; set; }
            //ΔFAM_RFU
            public decimal? lblORrfuA1 { get { return lblORrfu1 - lblsumRFCB; } }
            public decimal? lblORrfuA2 { get { return lblORrfu2 - lblsumRFCB; } }
            public decimal? lblORrfuA3 { get { return lblORrfu3 - lblsumRFCB; } }
            public decimal? lblORrfuA4 { get { return lblORrfu4 - lblsumRFCB; } }
            //ROX_CT
            public decimal? lblcq1 { get; set; }
            public decimal? lblcq2 { get; set; }
            public decimal? lblcq3 { get; set; }
            public decimal? lblcq4 { get; set; }
            //ΔROX_CT
            public decimal? lblctNA1 { get { return lblcq1 - lblsumctC; } }
            public decimal? lblctNA2 { get { return lblcq2 - lblsumctC; } }
            public decimal? lblctNA3 { get { return lblcq3 - lblsumctC; } }
            public decimal? lblctNA4 { get { return lblcq4 - lblsumctC; } }
            //ROX_RFU
            public decimal? lblNrfu1 { get; set; }
            public decimal? lblNrfu2 { get; set; }
            public decimal? lblNrfu3 { get; set; }
            public decimal? lblNrfu4 { get; set; }
            //ΔROX_RFU
            public decimal? lblNrfuA1 { get { return lblNrfu1 - lblsumRFCC; } }
            public decimal? lblNrfuA2 { get { return lblNrfu2 - lblsumRFCC; } }
            public decimal? lblNrfuA3 { get { return lblNrfu3 - lblsumRFCC; } }
            public decimal? lblNrfuA4 { get { return lblNrfu4 - lblsumRFCC; } }
            //阳性质控
            public decimal ?lblsumctA { get; set; }
            public decimal ?lblsumctB{ get; set; }
            public decimal ?lblsumctC{ get; set; }

            public decimal lblsumRFCA { get; set; }
            public decimal lblsumRFCB { get; set; }
            public decimal lblsumRFCC { get; set; }
           
            //结论
            public string Result1 { get; set; }
            public string Result2 { get; set; }
            public string Result3 { get; set; }
            public string Result4 { get; set; }
            //检查时间
            public string Datetime { get; set; }

        }

        private void 删除当天数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var sql1 = $"DELETE FROM CT_DB WHERE DATE='{DateTime.Now.ToString("yyyy-MM-dd")}'";
                var sql2 = $"DELETE FROM RFU_DB WHERE DATE='{DateTime.Now.ToString("yyyy-MM-dd")}'";
                SqLiteHelper.ExecuteSql(sql1);
                SqLiteHelper.ExecuteSql(sql2);
                MessageBox.Show("删除成功！","",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show($"删除失败，原因：{ex.Message}！", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void 删除所有数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var sql1 = $"DELETE FROM CT_DB";
                var sql2 = $"DELETE FROM RFU_DB";
                SqLiteHelper.ExecuteSql(sql1);
                SqLiteHelper.ExecuteSql(sql2);
                MessageBox.Show("删除成功！", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show($"删除失败，原因：{ex.Message}！", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
