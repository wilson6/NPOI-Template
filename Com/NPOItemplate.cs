using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.Data;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

 class NPOItemplate
    {

        /// <summary>
        /// 填充字典里面的字段 如 #key#  
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="Dic"></param>
        private static void _InitTmpFromDictionary(ISheet sheet, Dictionary<string, string> Dic)
        {
            Func<Match, string> initprams = (m =>
            {
                var res = string.Empty;
                if (m.Value.Contains("#{"))
                {
                    res = m.Value.Replace("#{", "#").Replace("}#", "#");
                }
                else
                {
                    var key = m.Value.TrimStart('#').TrimEnd('#');
                    res = Dic.ContainsKey(key) ? Dic[key] : "";
                }
                return res;
            });

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) { continue; }
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    var cellobj = row.GetCell(j);
                    if (cellobj != null)
                    {
                        var cellvalue = cellobj.ToString();
                        if (!string.IsNullOrEmpty(cellvalue)&&cellvalue.Contains("#"))
                        {
                            cellvalue = Regex.Replace(cellvalue, @"#[a-zA-Z_\{\}0-9]+?#", new MatchEvaluator(initprams), System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            cellobj.SetCellValue(cellvalue);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// [tmp]模版插入数据
        /// </summary>
        /// <param name="sheet">工作簿</param>
        /// <param name="detailtable">数据datatable</param>
        /// <param name="rowIndex">模版的行的index</param>
        private static void _InitTmpDetailFromDataTable(ISheet sheet, DataTable detailtable, ref int rowIndex)
        {
            if (detailtable.Rows.Count > 0)
            {
                var tmpindex = rowIndex;//记录模版的index，追加完成后删除
                var itemcount = 0;
                ICell datacell;
                IRow tmprow = sheet.GetRow(rowIndex);
                #region 委托
                DataRow rowfield = detailtable.Rows[0];
                Func<Match, string> initprams = (m =>
                {
                    var res = string.Empty;
                    if (m.Value.Contains("{"))
                    {
                        res = m.Value.Replace("{", "").Replace("}", "");
                    }
                    else
                    {
                        var key = m.Value.TrimStart('#').TrimEnd('#');
                        res = detailtable.Columns.Contains(key) && key != "" ? rowfield[key].ToString() : "";
                    }
                    return res;
                });
                #endregion
                 
                foreach (DataRow row in detailtable.Rows)
                {
                    itemcount++;
                    rowIndex++; 
                    
                    //这里有个bug 移动区域有合并列的时候 有可能报“索引超出范围。必须为非负值并小于集合大小”错误
                    sheet.ShiftRows(rowIndex, //开始行
                        sheet.LastRowNum, //结束行
                        1, //插入行总数
                        true,        //是否复制行高
                        false        //是否重置行高
                        );
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    #region 新追加的行 追加列 并填充数据
                    rowfield = row;
                    foreach (ICell col in tmprow)
                    {
                        datacell = dataRow.CreateCell(col.ColumnIndex);
                        datacell.CellStyle = col.CellStyle;
                        var field = col.ToString();
                        var cellvalue = col.ToString().Replace("[tmp]", "");
                        if (!string.IsNullOrEmpty(cellvalue) && cellvalue.Contains("#"))
                        {
                            cellvalue = Regex.Replace(cellvalue, @"#[a-zA-Z_\{\}]+?#", new MatchEvaluator(initprams), System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            datacell.SetCellValue(cellvalue);
                        }
                    }
                    #endregion
                }
                
                //清除模版行
                 sheet.ShiftRows(tmpindex+1, //开始行
                        sheet.LastRowNum, //结束行
                        -1
                        );
            }
        }

        /// <summary>
        /// 主要方法
        /// </summary>
        /// <param name="TemplateServerPath"></param>
        /// <param name="Dic"></param>
        /// <param name="detailtable"></param>
        /// <returns></returns>
        
        public static IWorkbook GenerateIWorkbook(string TemplateServerPath, Dictionary<string, string> Dic, DataTable detailtable)
        {

            IWorkbook hssfworkbook = null;
            string fileExt = "";
            MemoryStream ms = new MemoryStream();
            using (FileStream file = new FileStream(TemplateServerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                fileExt = Path.GetExtension(TemplateServerPath).ToLower();
                if (fileExt == ".xlsx")
                {
                    hssfworkbook = new XSSFWorkbook(file);
                }
                else if (fileExt == ".xls")
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }     
                file.Close();
            }
            if (hssfworkbook != null)
            {
                ISheet sheet = hssfworkbook.GetSheetAt(0);
                _InitTmpFromDictionary(sheet, Dic);
                var rowIndex = -1;
                for (int i = 0; i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) { continue; }
                    var cellobj = row.GetCell(0);
                    if (cellobj != null && cellobj.ToString().Contains("[tmp]"))
                    {
                        rowIndex = i;
                        break;
                    }
                }
                _InitTmpDetailFromDataTable(sheet,detailtable,ref rowIndex);
            }
            return hssfworkbook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static MemoryStream IworkToMemoryStream(IWorkbook workbook){
                 //hssfworkbook.Write(new FileStream(outpath+"out"+fileExt, FileMode.Create, FileAccess.Write,FileShare.ReadWrite));
                 MemoryStream ms = new MemoryStream();
                 workbook.Write(ms);                
                 ms.Flush();
                 return ms;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="outpath"></param>
        public static void IworkSave(IWorkbook workbook,string outpath){
                var filename=Path.GetFileName(outpath);
                if(!Directory.Exists(Path.GetDirectoryName(outpath))){
                    Directory.CreateDirectory(Path.GetDirectoryName(outpath));
                }
                workbook.Write(new FileStream(outpath, FileMode.Create, FileAccess.Write,FileShare.ReadWrite));
        }                
    }

