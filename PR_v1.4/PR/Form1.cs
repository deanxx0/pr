using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Reflection;

// git test test
namespace PR
{
    public partial class Form1 : Form
    {
        public class Schedule
        {
            public string strEquip;
            public int nProductNumber;
            public string strLotName;
            public int nStripCount;
            public string strState;
        }
        public class TheData
        {
            public string strDate;
            public string strEquip;
            public string strProductNumber;
            public string strLotName;
            public string strStripCount;

            public int nFL_RCSum;
            public int nBL_RCSum;
            public int nFL_NGSum;
            public int nBL_NGSum;
            public int nFA_OKSum;
            public int nBA_OKSum;
            public int nFA_NGSum;
            public int nBA_NGSum;
            public int nFA_RCSum;
            public int nBA_RCSum;

            public float fFL_TimeAvg;
            public float fFA_TimeAvg;
            public float fBL_TimeAvg;
            public float fBA_TimeAvg;
            public float fReview_TimeAvg;
        }

        public string _strDate;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dgvSet();
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            _strDate = monthCalendar1.SelectionStart.ToString("yyyyMMdd");
            string strD = monthCalendar1.SelectionStart.ToString("yyyy/MM/dd");
            laDate.Text = strD;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 50; 

            List<List<Schedule>> schedules = new List<List<Schedule>>();

            Task task = Task.Factory.StartNew(() =>
            {
                GetSchedules(_strDate, ref schedules);
                
                return schedules;
            }).ContinueWith(s => 
            {
                this.Invoke(new System.Action(delegate
                {
                    ShowSchedules(s.Result);
                    //ShowSchedules(schedules); //s.result랑 같다
                    this.progressBar1.Style = ProgressBarStyle.Blocks;
                    this.progressBar1.Value = 0;
                }));
            });
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            dgv.SelectAll();
        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            string path = "";
            FolderBrowserDialog f = new FolderBrowserDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                path = f.SelectedPath;
                tbxPath.Text = path;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //path check
            string path = tbxPath.Text;
            if(path == null || path == "")
            {
                MessageBox.Show("Enter the Path");
                return;
            }
            else
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(path);
                }
                catch
                {
                    MessageBox.Show("There is no Path");
                    return;
                }
            }

            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 50;

            Task task = Task.Factory.StartNew(() =>
            {
                List<TheData> theDatas = new List<TheData>();
                GetTheDatas(ref theDatas);

                List<TheData> lastDatas = new List<TheData>();
                GetLastDatas(theDatas, ref lastDatas);

                CreateExcel(lastDatas, path);

            }).ContinueWith(r =>
            {
                this.Invoke(new System.Action(delegate
                {
                    this.progressBar1.Style = ProgressBarStyle.Blocks;
                    this.progressBar1.Value = 0;
                }));
                MessageBox.Show("Create Report Completed");
            });
        }

        public void dgvSet()
        {
            dgv.ColumnCount = 5;
            dgv.Columns[0].Name = "Equip";
            dgv.Columns[1].Name = "Product Number";
            dgv.Columns[2].Name = "Lot Name";
            dgv.Columns[3].Name = "Strip Quantity";
            dgv.Columns[4].Name = "State";
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.RowHeadersVisible = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = true;
            dgv.AllowUserToAddRows = false;
        }

        public void GetSchedules(string strDate, ref List<List<Schedule>> schedules)
        {
            List<Schedule> schedule1 = new List<Schedule>();
            List<Schedule> schedule2 = new List<Schedule>();
            List<Schedule> schedule3 = new List<Schedule>();
            List<Schedule> schedule4 = new List<Schedule>();
            List<Schedule> schedule5 = new List<Schedule>();
            List<Schedule> schedule6 = new List<Schedule>();
            List<Schedule> schedule7 = new List<Schedule>();
            List<Schedule> schedule8 = new List<Schedule>();
            List<Schedule> schedule9 = new List<Schedule>();
            List<Schedule> schedule10 = new List<Schedule>();
            schedules.Add(schedule1);
            schedules.Add(schedule2);
            schedules.Add(schedule3);
            schedules.Add(schedule4);
            schedules.Add(schedule5);
            schedules.Add(schedule6);
            schedules.Add(schedule7);
            schedules.Add(schedule8);
            schedules.Add(schedule9);
            schedules.Add(schedule10);

            int equip = 1;
            string strDBConn;
            foreach(List<Schedule> s in schedules)
            {
                strDBConn = string.Format("server=192.168.118.{0}1;user=TEST;database=DB_A2_TEST;port=3306;password=1234;SslMode=None; Allow User Variables=True", equip);

                try
                {
                    using (MySqlConnection conn = new MySqlConnection(strDBConn))
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();

                        using (MySqlCommand cmd = new MySqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.CommandText = string.Format("SELECT PRODUCT_NO, LOT_NAME, WORK_STATUS FROM TBL_a2_S02_SCHEDULE WHERE PRODUCT_DATE = {0}", strDate);

                            MySqlDataReader reader = cmd.ExecuteReader();
                            while(reader.Read())
                            {
                                Schedule schedule = new Schedule();
                                schedule.strEquip = equip.ToString();
                                schedule.nProductNumber = reader.GetInt32("PRODUCT_NO");
                                schedule.strLotName = reader.GetString("LOT_NAME");
                                schedule.strState = reader.GetString("WORK_STATUS");

                                s.Add(schedule);
                            }
                            reader.Close();

                            for (int i = 0; i < s.Count; i++)
                            {
                                Schedule schedule = s[i];
                                //cmd.CommandText = string.Format("SELECT COUNT(*) FROM TBL_A2_S03_STRIP WHERE PRODUCT_DATE='{0}' AND PRODUCT_NO={1} AND STRIP_STATUS != 'XOUT'", strDate, schedule.nProductNumber);
                                cmd.CommandText = string.Format("SELECT COUNT(*) FROM TBL_A2_S03_STRIP WHERE PRODUCT_DATE='{0}' AND PRODUCT_NO={1}", strDate, schedule.nProductNumber);

                                long fStripCount = (long)cmd.ExecuteScalar();
                                s[i].nStripCount = (int)fStripCount;
                            }
                        }
                    }
                }
                catch(Exception)
                {
                    Console.WriteLine("GetSchedules error");
                }

                equip++;
            }
        }

        public void ShowSchedules(List<List<Schedule>> schedules)
        {
            dgv.Rows.Clear();

            foreach(List<Schedule> s in schedules)
            {
                for (int i = 0; i < s.Count; i++)
                {
                    dgv.Rows.Add(s[i].strEquip, s[i].nProductNumber, s[i].strLotName, s[i].nStripCount, s[i].strState);
                }
            }
        }
        public void GetTheDatas(ref List<TheData> theDatas)
        {
            string strEquip = "";
            string strProductNumber = "";

            for (int i = 0; i < dgv.SelectedRows.Count; i++)
            {
                strEquip = dgv.SelectedRows[i].Cells[0].Value.ToString();
                strProductNumber = dgv.SelectedRows[i].Cells[1].Value.ToString();

                TheData theData = new TheData();
                string strDBConn = string.Format("server=192.168.118.{0}1;user=TEST;database=DB_A2_TEST;port=3306;password=1234;SslMode=None; Allow User Variables=True", strEquip);

                try
                {
                    using (MySqlConnection conn = new MySqlConnection(strDBConn))
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();

                        using (MySqlCommand cmd = new MySqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.CommandText = String.Format(
                          @"SET @PRODUCT_DATE='{0}';
                            SET @PRODUCT_NO = {1};
                            SET @STRIP_NO = 0;

                            SELECT COUNT(A.STRIP_NO), SUM(A.FLS_RC), SUM(A.FLS_NG), SUM(A.FAS_OK), SUM(A.FAS_NG), SUM(A.FAS_RC), SUM(A.BLS_RC), SUM(A.BLS_NG), SUM(A.BAS_OK), SUM(A.BAS_NG), SUM(A.BAS_RC),
                                    AVG(TB1.FLS_TIME), AVG(TB1.FAS_TIME), AVG(TB1.BLS_TIME), AVG(TB1.BAS_TIME), AVG(TB1.RV_TIME)
                        
                            FROM(SELECT STRIP_NO,
                                SUM(case when(RST_CODE_0 = 'NG' and FB = 'F' and unit_id != -1) then 1 else 0 end) as FLS_NG,
                                SUM(case when(RST_CODE_0 = 'RC' and FB = 'F' and unit_id != -1) then 1 else 0 end) as FLS_RC,
                                SUM(case when(RST_CODE_1 = 'OK' and FB = 'F') then 1 else 0 end) as FAS_OK,
                                SUM(case when(RST_CODE_1 = 'RC' and FB = 'F') then 1 else 0 end) as FAS_RC,
                                SUM(case when(RST_CODE_1 = 'NG' and FB = 'F') then 1 else 0 end) as FAS_NG,

                                SUM(case when(RST_CODE_0 = 'NG' and FB = 'B' and unit_id != -1) then 1 else 0 end) as BLS_NG,
                                SUM(case when(RST_CODE_0 = 'RC' and FB = 'B' and unit_id != -1) then 1 else 0 end) as BLS_RC,
                                SUM(case when(RST_CODE_1 = 'OK' and FB = 'B') then 1 else 0 end) as BAS_OK,
                                SUM(case when(RST_CODE_1 = 'NG' and FB = 'B') then 1 else 0 end) as BAS_NG,
                                SUM(case when(RST_CODE_1 = 'RC' and FB = 'B') then 1 else 0 end) as BAS_RC
                        
                            FROM TBL_A2_R02_DEFECT_INFO             
                            WHERE PRODUCT_DATE = @PRODUCT_DATE and PRODUCT_NO = @PRODUCT_NO and STRIP_NO >= @STRIP_NO
                            GROUP BY STRIP_NO) AS A
                        
                            RIGHT JOIN(SELECT B.STRIP_NO, B.FLS_TIME, B.FAS_TIME, B.BLS_TIME, B.BAS_TIME, B.RV_TIME
                        
                                FROM(SELECT STRIP_NO,
                                    TIMESTAMPDIFF(SECOND, FLS_START_DATE, FLS_END_DATE) as FLS_TIME,
                                    TIMESTAMPDIFF(SECOND, FAS_START_DATE, FAS_END_DATE) as FAS_TIME,
                                    TIMESTAMPDIFF(SECOND, BLS_START_DATE, BLS_END_DATE) as BLS_TIME,
                                    TIMESTAMPDIFF(SECOND, BAS_START_DATE, BAS_END_DATE) as BAS_TIME,
                                    TIMESTAMPDIFF(SECOND, REVIEW_START_DATE, REVIEW_END_DATE) as RV_TIME,
                                    REVIEW_END_DATE

                             FROM TBL_A2_S03_STRIP
                             WHERE PRODUCT_DATE = @PRODUCT_DATE AND PRODUCT_NO = @PRODUCT_NO AND STRIP_NO >= @STRIP_NO AND REVIEW_END_DATE IS NOT NULL
                             GROUP BY STRIP_NO) AS B) AS TB1
                             ON A.STRIP_NO = TB1.STRIP_NO", _strDate, strProductNumber);

                            MySqlDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                theData.strDate = _strDate;
                                theData.strEquip = strEquip;
                                theData.strProductNumber = strProductNumber;
                                theData.strStripCount = reader.GetInt32("COUNT(A.STRIP_NO)").ToString();

                                theData.nFL_RCSum = reader.GetInt32("SUM(A.FLS_RC)");
                                theData.nBL_RCSum = reader.GetInt32("SUM(A.BLS_RC)");
                                theData.nFL_NGSum = reader.GetInt32("SUM(A.FLS_NG)");
                                theData.nBL_NGSum = reader.GetInt32("SUM(A.BLS_NG)");
                                theData.nFA_OKSum = reader.GetInt32("SUM(A.FAS_OK)");
                                theData.nBA_OKSum = reader.GetInt32("SUM(A.BAS_OK)");
                                theData.nFA_NGSum = reader.GetInt32("SUM(A.FAS_NG)");
                                theData.nBA_NGSum = reader.GetInt32("SUM(A.BAS_NG)");
                                theData.nFA_RCSum = reader.GetInt32("SUM(A.FAS_RC)");
                                theData.nBA_RCSum = reader.GetInt32("SUM(A.BAS_RC)");

                                theData.fFL_TimeAvg = reader.GetFloat("AVG(TB1.FLS_TIME)");
                                theData.fFA_TimeAvg = reader.GetFloat("AVG(TB1.FAS_TIME)");
                                theData.fBL_TimeAvg = reader.GetFloat("AVG(TB1.BLS_TIME)");
                                theData.fBA_TimeAvg = reader.GetFloat("AVG(TB1.BAS_TIME)");
                                theData.fReview_TimeAvg = reader.GetFloat("AVG(TB1.RV_TIME)");
                            }
                            reader.Close();

                            //LotName query
                            cmd.CommandText = String.Format(@"
                            SET @PRODUCT_DATE='{0}';
                            SET @PRODUCT_NO = {1};
                            SET @STRIP_NO = 0;

                            Select LOT_NAME
                            From tbl_a2_s02_schedule
                            WHERE PRODUCT_DATE = @PRODUCT_DATE and PRODUCT_NO = @PRODUCT_NO", _strDate, strProductNumber);
                            reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                theData.strLotName = reader.GetString("LOT_NAME");
                            }
                            reader.Close();

                            theDatas.Add(theData);
                        }
                    }
                }
                catch(Exception)
                {
                    Console.WriteLine("GetTheDatas error");
                }
            }
        }

        public void GetLastDatas(List<TheData> theDatas, ref List<TheData> lastDatas)
        {
            List<string> duplicateLots = new List<string>();
            bool duplicateExist = false;

            HashSet<string> hashSet = new HashSet<string>();
            foreach(TheData d in theDatas)
            {
                if (hashSet.Add(d.strLotName))
                {
                }
                else //중복된 LotName이 들어온 경우
                {
                    //일단 중복된 LotName List를 얻어두자
                    duplicateLots.Add(d.strLotName);
                    duplicateExist = true;
                }
            }

            List<TheData> mergedDatas = new List<TheData>();
            
            //중복Lot 있는 경우
            if (duplicateExist == true)
            {
                List<int> dupIndexinTheDatas = new List<int>();
                //theDatas 에서 duplicateLots 에 포함된 이름이 잇는지 검사
                for (int i = 0; i < duplicateLots.Count; i++)
                {
                    TheData mergedData = new TheData();
                    List<TheData> dupDatas = new List<TheData>();

                    for (int n = 0; n < theDatas.Count; n++)
                    {
                        if(theDatas[n].strLotName == duplicateLots[i])
                        {
                            dupDatas.Add(theDatas[n]); //포함된게 있으면 dupDatas 에 따로 저장
                            dupIndexinTheDatas.Add(n); //나중에 theDatas에서 지울거니까 인덱스 저장해두기
                        }
                    }

                    //중복된 dupDatas를 이용해 하나의 mergedData로 병합
                    mergedData.strDate = dupDatas[0].strDate;
                    mergedData.strEquip = dupDatas[0].strEquip;
                    mergedData.strProductNumber = dupDatas[0].strProductNumber;
                    mergedData.strLotName = dupDatas[0].strLotName;

                    int total_StripCount = 0;
                    TheData dataBox = new TheData();
                    for (int n = 0; n < dupDatas.Count; n++)
                    {
                        total_StripCount += int.Parse(dupDatas[n].strStripCount);

                        dataBox.nFL_RCSum += dupDatas[n].nFL_RCSum;
                        dataBox.nBL_RCSum += dupDatas[n].nBL_RCSum;
                        dataBox.nFL_NGSum += dupDatas[n].nFL_NGSum;
                        dataBox.nBL_NGSum += dupDatas[n].nBL_NGSum;
                        dataBox.nFA_OKSum += dupDatas[n].nFA_OKSum;
                        dataBox.nBA_OKSum += dupDatas[n].nBA_OKSum;
                        dataBox.nFA_NGSum += dupDatas[n].nFA_NGSum;
                        dataBox.nBA_NGSum += dupDatas[n].nBA_NGSum;
                        dataBox.nFA_RCSum += dupDatas[n].nFA_RCSum;
                        dataBox.nBA_RCSum += dupDatas[n].nBA_RCSum;

                        dataBox.fFL_TimeAvg += dupDatas[n].fFL_TimeAvg * int.Parse(dupDatas[n].strStripCount);
                        dataBox.fFA_TimeAvg += dupDatas[n].fFA_TimeAvg * int.Parse(dupDatas[n].strStripCount);
                        dataBox.fBL_TimeAvg += dupDatas[n].fBL_TimeAvg * int.Parse(dupDatas[n].strStripCount);
                        dataBox.fBA_TimeAvg += dupDatas[n].fBA_TimeAvg * int.Parse(dupDatas[n].strStripCount);
                        dataBox.fReview_TimeAvg += dupDatas[n].fReview_TimeAvg * int.Parse(dupDatas[n].strStripCount);
                    }

                    mergedData.strStripCount = total_StripCount.ToString();

                    mergedData.nFL_NGSum = dataBox.nFL_NGSum;
                    mergedData.nFL_RCSum = dataBox.nFL_RCSum;
                    mergedData.nFA_OKSum = dataBox.nFA_OKSum;
                    mergedData.nFA_RCSum = dataBox.nFA_RCSum;
                    mergedData.nFA_NGSum = dataBox.nFA_NGSum;

                    mergedData.nBL_NGSum = dataBox.nBL_NGSum;
                    mergedData.nBL_RCSum = dataBox.nBL_RCSum;
                    mergedData.nBA_OKSum = dataBox.nBA_OKSum;
                    mergedData.nBA_RCSum = dataBox.nBA_RCSum;
                    mergedData.nBA_NGSum = dataBox.nBA_NGSum;

                    mergedData.fFL_TimeAvg = dataBox.fFL_TimeAvg / total_StripCount;
                    mergedData.fFA_TimeAvg = dataBox.fFA_TimeAvg / total_StripCount;
                    mergedData.fBL_TimeAvg = dataBox.fBL_TimeAvg / total_StripCount;
                    mergedData.fBA_TimeAvg = dataBox.fBA_TimeAvg / total_StripCount;
                    mergedData.fReview_TimeAvg = dataBox.fReview_TimeAvg / total_StripCount;

                    //mergedDatas 에 병합된 하나의 Lotname을 가진 Data Add
                    mergedDatas.Add(mergedData);
                }
                //theDatas 에서 중복되어 병합된 것들 삭제
                for(int n = dupIndexinTheDatas.Count - 1; n >= 0; n--)
                {
                    theDatas.RemoveAt(dupIndexinTheDatas[n]);
                }
                //중복이 삭제된 theDatas와 중복끼리 병합한 mergedDatas를 하나로 합침
                //theDatas + mergedDatas
                //최종 lastDatas 획득
                lastDatas.AddRange(theDatas);
                lastDatas.AddRange(mergedDatas);
            }
            else
            {
                //중복된게 없으면 그냥 theDatas가 최종 lastDatas
                lastDatas = theDatas;
            }
        }

        public void CreateExcel(List<TheData> lastDatas, string path)
        {
            Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook;
            try
            {
                workbook = exApp.Workbooks.Open(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", @"Template.xlsx"));
            }
            catch(Exception)
            {
                MessageBox.Show("There is no Template.xlsx");
                return;
            }
            
            Worksheet worksheet = workbook.Worksheets["Chart"];

            worksheet.Cells[2, 2] = _strDate;

            for (int i = 0; i < lastDatas.Count; i++)
            {
                worksheet.Cells[i + 6, 1] = lastDatas[i].strEquip;
                worksheet.Cells[i + 6, 2] = lastDatas[i].strProductNumber;
                worksheet.Cells[i + 6, 3] = lastDatas[i].strLotName;
                worksheet.Cells[i + 6, 4] = lastDatas[i].strStripCount;
                worksheet.Cells[i + 6, 5] = lastDatas[i].nFL_NGSum;
                worksheet.Cells[i + 6, 6] = lastDatas[i].nFL_RCSum;
                worksheet.Cells[i + 6, 7] = lastDatas[i].nFA_OKSum;
                worksheet.Cells[i + 6, 8] = lastDatas[i].nFA_NGSum;
                worksheet.Cells[i + 6, 9] = lastDatas[i].nFA_RCSum;
                worksheet.Cells[i + 6, 10] = lastDatas[i].nBL_NGSum;
                worksheet.Cells[i + 6, 11] = lastDatas[i].nBL_RCSum;
                worksheet.Cells[i + 6, 12] = lastDatas[i].nBA_OKSum;
                worksheet.Cells[i + 6, 13] = lastDatas[i].nBA_NGSum;
                worksheet.Cells[i + 6, 14] = lastDatas[i].nBA_RCSum;
                worksheet.Cells[i + 6, 20] = lastDatas[i].fFL_TimeAvg;
                worksheet.Cells[i + 6, 21] = lastDatas[i].fFA_TimeAvg;
                worksheet.Cells[i + 6, 22] = lastDatas[i].fBL_TimeAvg;
                worksheet.Cells[i + 6, 23] = lastDatas[i].fBA_TimeAvg;
                worksheet.Cells[i + 6, 24] = lastDatas[i].fReview_TimeAvg;
            }

            workbook.SaveAs(path + "\\" + _strDate + "_Report.xlsx");
            workbook.Close();
            exApp.Quit();
        }
    }
}
