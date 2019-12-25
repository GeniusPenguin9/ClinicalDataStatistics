using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClinicalDataStatistics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string RootDirectory;
        private List<string> filelist = new List<string>();
        /* ******************************************************************************************************************************** */
        private static float max_power = 15.0F;
        private static int max_temp = 100;
        //电极1~6对应各参数的列
        private static int col_time1 = 16;
        private static int col_temp1 = 18;  
        private static int col_power1 = 19;

        private static int col_time2 = 24;
        private static int col_temp2 = 26;
        private static int col_power2 = 27;

        private static int col_time3 = 32;
        private static int col_temp3 = 34;
        private static int col_power3 = 35;

        private static int col_time4 = 40;
        private static int col_temp4 = 42;
        private static int col_power4 = 43;

        private static int col_time5 = 48;
        private static int col_temp5 = 50;
        private static int col_power5 = 51;

        private static int col_time6 = 56;
        private static int col_temp6 = 58;
        private static int col_power6 = 59;

        //各时间点对应的行
        private static int row_time0 = 1;
        private static int row_time5 = 6;
        private static int row_time39 = 40;
        private static int row_time40 = 41;
        private static int row_time119 = 120;
        /* ******************************************************************************************************************************** */

        private void ChooseRootDirectory_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                RootDirectory = dialog.SelectedPath;
            }
        }

        private void CreateAnalysis_Click(object sender, EventArgs e)
        {
            FindTextFile(new DirectoryInfo(RootDirectory));

            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Add(true);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1] = "File Full Name";
            worksheet.Cells[1, 2] = "Channel";
            worksheet.Cells[1, 3] = "Joule";

            int row = 2;
            foreach (var fullname in filelist)
            {
                if (fullname.EndsWith("txt"))
                {
                    worksheet.Cells[row, 1] = fullname;

                    var fileStream = new FileStream(fullname, FileMode.Open, FileAccess.Read);
                    using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                    {
                        string txtLine;
                        
                        int channel = 9;
                        float sumjoule = 0;

                        while ((txtLine = streamReader.ReadLine()) != null)
                        {
                            if (txtLine.Equals("") || Regex.IsMatch(txtLine, @"^[^0-9].+"))
                                continue;
                            else
                            {
                                var txtlinearray = txtLine.Split(' ');
                                if (txtlinearray.Length == 6)
                                {
                                    if(int.Parse(txtlinearray[2]) != 0)
                                        channel = int.Parse(txtlinearray[1]);
                                    sumjoule += float.Parse(txtlinearray[3]);
                                } 
                                
                            }
                        }
                        worksheet.Cells[row, 2] = channel;
                        worksheet.Cells[row, 3] = sumjoule;
                        row++;
                    }
                }
            }
        }
        public void FindTextFile(FileSystemInfo info)
        {
            if (!info.Exists) return;
            DirectoryInfo dir = info as DirectoryInfo;

            //不是目录 
            if (dir == null) return;
            FileSystemInfo[] files = dir.GetFileSystemInfos();
            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = files[i] as FileInfo;
                //是文件 
                if (file != null)
                {
                    filelist.Add(file.FullName);
                }

                //对于子目录，进行递归调用 
                else
                    FindTextFile(files[i]);
            }
        }

        private void CreateExcelAnalysis_Click(object sender, EventArgs e)
        {
            FindTextFile(new DirectoryInfo(RootDirectory));

            Excel.Application app = new Excel.Application();
            app.Visible = true;
            //app.UserControl = true;


            foreach (var fullname in filelist)
            {
                if (fullname.EndsWith(".xls"))
                {
                    var fullname_list = new List<string>(fullname.Split('\\'));
                    if (fullname_list.Last().Count() > 10)
                    {
                        //Excel.Workbook workbook = app.Workbooks.Open(fullname, 1,
                        //                                                false, 6, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                        //                                                "\t", false, false, 0, true, 1, 0);
                        Excel.Workbook workbook = app.Workbooks.Open(fullname);
                        Excel.Worksheet worksheet = workbook.ActiveSheet;

                        worksheet.Cells[239, 41] = "能量（焦耳）";//AO239

                        //合并单元格
                        Excel.Range range = worksheet.Range[worksheet.Cells[239, 41], worksheet.Cells[239, 43]];
                        range.Merge(true);

                        worksheet.Cells[240, 41] = "MAX";//AO240
                        worksheet.Cells[240, 42] = "AVG";//AP240
                        worksheet.Cells[240, 43] = "MIN";//AQ240

                        worksheet.Cells[241, 41].Formula = "=SUM(S1:S120)";//AO241
                        worksheet.Cells[241, 42].Formula = "=AVERAGE(S1:S120)";//AP241
                        worksheet.Cells[241, 43].Formula = "=S1";//AQ241

                        worksheet.Cells[242, 41].Formula = "=SUM(AA1:AA120)";//AO242
                        worksheet.Cells[242, 42].Formula = "=AVERAGE(AA1:AA120)";//AP242
                        worksheet.Cells[242, 43].Formula = "=AA1";//AQ242

                        worksheet.Cells[243, 41].Formula = "=SUM(AI1:AI120)";//AO243
                        worksheet.Cells[243, 42].Formula = "=AVERAGE(AI1:AI120)";//AP243
                        worksheet.Cells[243, 43].Formula = "=AI1";//AQ243

                        worksheet.Cells[244, 41].Formula = "=SUM(AQ1:AQ120)";//AO244
                        worksheet.Cells[244, 42].Formula = "=AVERAGE(AQ1:AQ120)";//AP244
                        worksheet.Cells[244, 43].Formula = "=AQ1";//AQ244

                        worksheet.Cells[245, 41].Formula = "=SUM(AY1:AY120)";//AO245
                        worksheet.Cells[245, 42].Formula = "=AVERAGE(AY1:AY120)";//AP245
                        worksheet.Cells[245, 43].Formula = "=AY1";//AQ245

                        worksheet.Cells[246, 41].Formula = "=SUM(BG1:BG120)";//AO246
                        worksheet.Cells[246, 42].Formula = "=AVERAGE(BG1:BG120)";//AP246
                        worksheet.Cells[246, 43].Formula = "=BG1";//AQ246

                        
                        string spath = @"E:\ClientD\" + fullname_list[4] + @"\";
                        if (!Directory.Exists(spath))
                            Directory.CreateDirectory(spath);

                        workbook.SaveAs(spath + fullname_list.Last());
                        workbook.Close();
                    }
                }
                

            }
            app.Quit();
        }

        private void CreateStatistics_Click(object sender, EventArgs e)
        {
            
            
            FindTextFile(new DirectoryInfo(RootDirectory));

            Excel.Application app = new Excel.Application();
            app.Visible = true;

            Excel.Workbook wb_new = app.Workbooks.Add(true);
            Excel.Worksheet ws_new = wb_new.ActiveSheet;
            CopyNewTile(ws_new);
            int row_new = 2;

            //fullname example：　E:\[0]Work\数据统计\原始数据整理文档\***\***＼文件名.xls
            //                    0  1       2       3               4                      
            foreach (var fullname in filelist)
            {
                if (fullname.EndsWith(".xls"))
                {
                    var fullname_list = new List<string>(fullname.Split('\\'));
                    if (fullname_list.Last().Count() > 10)
                    {
                        Excel.Workbook wb_org = app.Workbooks.Open(fullname);
                        Excel.Worksheet ws_org = wb_org.Worksheets["Sheet1"];

                        ws_new.Cells[row_new, 1] = fullname_list[4];
                        ws_new.Cells[row_new, 2] = GetLocation(fullname_list);
                        ws_new.Cells[row_new, 3] = "6";

                        ws_new.Cells[row_new, 4] = CalculateWorkPoint(ws_org);
                        int active_work_point = CalculateActiveWorkPoint(ws_org);
                        ws_new.Cells[row_new, 5] = active_work_point;
                        ws_new.Cells[row_new, 6] = (float)active_work_point / 6;

                        float whole_power = CalculateWholePower(ws_org);
                        float active_power = CalculateActivePower(ws_org);
                        ws_new.Cells[row_new, 8] = whole_power;
                        ws_new.Cells[row_new, 9] = active_power;
                        ws_new.Cells[row_new, 10] = active_power / whole_power;

                        var time = CalculateTime(ws_org);   //[0] whole time; [1] active time
                        ws_new.Cells[row_new, 12] = time[0];
                        ws_new.Cells[row_new, 13] = time[1];
                        ws_new.Cells[row_new, 14] = (float) time[1] / (float)time[0];

                        wb_org.Close();
                        row_new += 1;
                    }

                }
            }
        }

        private void CopyNewTile(Excel.Worksheet ws)
        {
            ws.Cells[1, 1] = "病例";
            ws.Cells[1, 2] = "部位";
            ws.Cells[1, 3] = "总电极数量";
            ws.Cells[1, 4] = "工作电极数量";
            ws.Cells[1, 5] = "有效电极数量";
            ws.Cells[1, 6] = "电极有效比";
            //ws.Cells[1, 7] = "部位";
            ws.Cells[1, 8] = "总输出功率（瓦秒）";
            ws.Cells[1, 9] = "有效输出功率（瓦秒）";
            ws.Cells[1, 10] = "功率有效比";
            //ws.Cells[1, 11] = "部位";
            ws.Cells[1, 12] = "总输出时间（秒）";
            ws.Cells[1, 13] = "有效输出时间（秒）";
            ws.Cells[1, 14] = "时间有效比";


            ws.Cells[1, 15] = "总电极数量";
            ws.Cells[1, 16] = "工作电极数量";
            ws.Cells[1, 17] = "有效电极数量";
            ws.Cells[1, 18] = "电极有效比";
            //ws.Cells[1, 19] = "部位";
            ws.Cells[1, 20] = "总输出功率（瓦秒）";
            ws.Cells[1, 21] = "有效输出功率（瓦秒）";
            ws.Cells[1, 22] = "功率有效比";
            //ws.Cells[1, 23] = "部位";
            ws.Cells[1, 24] = "总输出时间（秒）";
            ws.Cells[1, 25] = "有效输出时间（秒）";
            ws.Cells[1, 26] = "时间有效比";
        }

        private string GetLocation(List<string> fullname_list)
        {
            string res = "";
            for(var i = 5; i < fullname_list.Count() - 1; i++)
            {
                res += fullname_list[i];
            }
            return res;
        }

        private int CalculateWorkPoint(Excel.Worksheet ws_org)
        {
            int res = 0;
            if (int.Parse(ws_org.Cells[row_time5, col_temp1].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp1].Value.ToString()) < max_temp)
                res += 1;
            if (int.Parse(ws_org.Cells[row_time5, col_temp2].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp2].Value.ToString()) < max_temp)
                res += 1;
            if (int.Parse(ws_org.Cells[row_time5, col_temp3].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp3].Value.ToString()) < max_temp)
                res += 1;
            if (int.Parse(ws_org.Cells[row_time5, col_temp4].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp4].Value.ToString()) < max_temp)
                res += 1;
            if (int.Parse(ws_org.Cells[row_time5, col_temp5].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp5].Value.ToString()) < max_temp)
                res += 1;
            if (int.Parse(ws_org.Cells[row_time5, col_temp6].Value.ToString()) > 0 &&
                int.Parse(ws_org.Cells[row_time5, col_temp6].Value.ToString()) < max_temp)
                res += 1;

            return res;
        }

        private int CalculateActiveWorkPoint(Excel.Worksheet ws_org)
        {
            int res = 0;
            
            var workarea_temp1 = ws_org.Range[ws_org.Cells[row_time40, col_temp1], ws_org.Cells[row_time119, col_temp1]].Value2;
            var workarea_temp2 = ws_org.Range[ws_org.Cells[row_time40, col_temp2], ws_org.Cells[row_time119, col_temp2]].Value2;
            var workarea_temp3 = ws_org.Range[ws_org.Cells[row_time40, col_temp3], ws_org.Cells[row_time119, col_temp3]].Value2;
            var workarea_temp4 = ws_org.Range[ws_org.Cells[row_time40, col_temp4], ws_org.Cells[row_time119, col_temp4]].Value2;
            var workarea_temp5 = ws_org.Range[ws_org.Cells[row_time40, col_temp5], ws_org.Cells[row_time119, col_temp5]].Value2;
            var workarea_temp6 = ws_org.Range[ws_org.Cells[row_time40, col_temp6], ws_org.Cells[row_time119, col_temp6]].Value2;

            if (IsActiveWorkPoint(workarea_temp1))
                res += 1;
            if (IsActiveWorkPoint(workarea_temp2))
                res += 1;
            if (IsActiveWorkPoint(workarea_temp3))
                res += 1;
            if (IsActiveWorkPoint(workarea_temp4))
                res += 1;
            if (IsActiveWorkPoint(workarea_temp5))
                res += 1;
            if (IsActiveWorkPoint(workarea_temp6))
                res += 1;

            return res;
        }

        private bool IsActiveWorkPoint(dynamic workarea_temp)
        {
            var whole = 0;
            var active = 0;

            for (var i = 1; i <= row_time119 - row_time40 + 1; i++)//value2 row start from 1
            {
                if (workarea_temp[i, 1] == 0)
                    break;
                else if (workarea_temp[i, 1] >= 58 && workarea_temp[i, 1] < max_temp)
                {
                    active += 1;
                    whole += 1;
                }
                else
                    whole += 1;
            }
            if (whole == 0)
                return false;   
            return (float)active / (float)whole >= 0.9 ? true : false;
        }

        private float CalculateWholePower(Excel.Worksheet ws_org)
        {
            float res = 0;

            var workarea_power1 = ws_org.Range[ws_org.Cells[row_time0, col_power1], ws_org.Cells[row_time119, col_power1]].Value2;
            var workarea_power2 = ws_org.Range[ws_org.Cells[row_time0, col_power2], ws_org.Cells[row_time119, col_power2]].Value2;
            var workarea_power3 = ws_org.Range[ws_org.Cells[row_time0, col_power3], ws_org.Cells[row_time119, col_power3]].Value2;
            var workarea_power4 = ws_org.Range[ws_org.Cells[row_time0, col_power4], ws_org.Cells[row_time119, col_power4]].Value2;
            var workarea_power5 = ws_org.Range[ws_org.Cells[row_time0, col_power5], ws_org.Cells[row_time119, col_power5]].Value2;
            var workarea_power6 = ws_org.Range[ws_org.Cells[row_time0, col_power6], ws_org.Cells[row_time119, col_power6]].Value2;
            foreach (var power in workarea_power1)
                res += (power > 0 && power < max_power) ? power : 0;
            foreach (var power in workarea_power2)
                res += (power > 0 && power < max_power) ? power : 0;
            foreach (var power in workarea_power3)
                res += (power > 0 && power < max_power) ? power : 0;
            foreach (var power in workarea_power4)
                res += (power > 0 && power < max_power) ? power : 0;
            foreach (var power in workarea_power5)
                res += (power > 0 && power < max_power) ? power : 0;
            foreach (var power in workarea_power6)
                res += (power > 0 && power < max_power) ? power : 0;

            return res;
        }

        private float CalculateActivePower(Excel.Worksheet ws_org)
        {
            float res = 0;

            var workarea_temp_power1 = ws_org.Range[ws_org.Cells[row_time0, col_temp1], ws_org.Cells[row_time119, col_power1]].Value2;
            var workarea_temp_power2 = ws_org.Range[ws_org.Cells[row_time0, col_temp2], ws_org.Cells[row_time119, col_power2]].Value2;
            var workarea_temp_power3 = ws_org.Range[ws_org.Cells[row_time0, col_temp3], ws_org.Cells[row_time119, col_power3]].Value2;
            var workarea_temp_power4 = ws_org.Range[ws_org.Cells[row_time0, col_temp4], ws_org.Cells[row_time119, col_power4]].Value2;
            var workarea_temp_power5 = ws_org.Range[ws_org.Cells[row_time0, col_temp5], ws_org.Cells[row_time119, col_power5]].Value2;
            var workarea_temp_power6 = ws_org.Range[ws_org.Cells[row_time0, col_temp6], ws_org.Cells[row_time119, col_power6]].Value2;

            res += CaluculateActivePowerOneCh(workarea_temp_power1);
            res += CaluculateActivePowerOneCh(workarea_temp_power2);
            res += CaluculateActivePowerOneCh(workarea_temp_power3);
            res += CaluculateActivePowerOneCh(workarea_temp_power4);
            res += CaluculateActivePowerOneCh(workarea_temp_power5);
            res += CaluculateActivePowerOneCh(workarea_temp_power6);

            return res;
        }

        private float CaluculateActivePowerOneCh (dynamic workarea_temp_power)
        {
            float res = 0;

            for(var i = 1; i <= row_time119 - row_time0 + 1; i++)
            {
                if (workarea_temp_power[i, 1] >= 58 && workarea_temp_power[i, 1] < max_temp)
                    res += workarea_temp_power[i, 2];
            }
            return res;
        }

        private List<int> CalculateTime (Excel.Worksheet ws_org)
        {
            List<int> res = new List<int> {0, 0};

            var workarea_temp1 = ws_org.Range[ws_org.Cells[row_time0, col_temp1], ws_org.Cells[row_time119, col_temp1]].Value2;
            var workarea_temp2 = ws_org.Range[ws_org.Cells[row_time0, col_temp2], ws_org.Cells[row_time119, col_temp2]].Value2;
            var workarea_temp3 = ws_org.Range[ws_org.Cells[row_time0, col_temp3], ws_org.Cells[row_time119, col_temp3]].Value2;
            var workarea_temp4 = ws_org.Range[ws_org.Cells[row_time0, col_temp4], ws_org.Cells[row_time119, col_temp4]].Value2;
            var workarea_temp5 = ws_org.Range[ws_org.Cells[row_time0, col_temp5], ws_org.Cells[row_time119, col_temp5]].Value2;
            var workarea_temp6 = ws_org.Range[ws_org.Cells[row_time0, col_temp6], ws_org.Cells[row_time119, col_temp6]].Value2;

            var res1 = CalculateTimeOneCh(workarea_temp1);
            var res2 = CalculateTimeOneCh(workarea_temp2);
            var res3 = CalculateTimeOneCh(workarea_temp3);
            var res4 = CalculateTimeOneCh(workarea_temp4);
            var res5 = CalculateTimeOneCh(workarea_temp5);
            var res6 = CalculateTimeOneCh(workarea_temp6);

            res[0] = res1[0] + res2[0] + res3[0] + res4[0] + res5[0] + res6[0];
            res[1] = res1[1] + res2[1] + res3[1] + res4[1] + res5[1] + res6[1];

            return res;//[0] whole time; [1] active time
        }

        private List<int> CalculateTimeOneCh (dynamic workearea_temp)
        {
            List<int> res = new List<int> { 0, 0};
            //[0] whole time; [1] active time
            for(var i = 1; i <= row_time119 - row_time0 + 1; i++)
            {
                if(workearea_temp[i, 1] > 0 && workearea_temp[i, 1] < max_temp)
                {
                    res[0] += 1;
                    if (workearea_temp[i, 1] >= 58 && workearea_temp[i, 1] < max_temp)
                        res[1] += 1;
                }
            }
            return res;
        }
    }

    public class TxtLine
    {
        private int time;
        private int channel;
        private int temperature;
        private float power;
        private int resistance;
        private int temp_setting;
        public TxtLine(string txtline)
        {
            var txt = txtline.Split(' ');
            this.time = int.Parse(txt[0]);
            this.channel = int.Parse(txt[1]);
            this.temperature = int.Parse(txt[2]);
            this.power = float.Parse(txt[3]);
            this.resistance = int.Parse(txt[4]);
            this.temp_setting = int.Parse(txt[5]);
        }
    }
}


