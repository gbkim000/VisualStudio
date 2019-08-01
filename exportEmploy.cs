using System;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;

namespace WindowsFormsApplication1
{
    public partial class exportEmploy : Form
    {
        //string dbFile = login.dbFile;
        string dbFile = "E:\\VS2019\\Data\\취업관리2.db";
        class Company
        {
            public string Code;
            public string Name;
        }
        List<Company> compList = new List<Company>();

        public exportEmploy()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;

            dataGridView1.Columns.Add("ID", "학 번");
            dataGridView1.Columns.Add("name", "이    름");
            dataGridView1.Columns.Add("sex", "성별");
            dataGridView1.Columns.Add("CODE", "코 드");
            dataGridView1.Columns.Add("company", "기 업 체 명");
            dataGridView1.Columns.Add("manager", "담당자");
            dataGridView1.Columns.Add("phone", "전화번호");
            dataGridView1.Columns.Add("salary", "월급여");
            dataGridView1.Columns.Add("jdate", "취업일자");
            dataGridView1.Columns.Add("udate", "진학일자");
            dataGridView1.Columns.Add("history1", "변동1");
            dataGridView1.Columns.Add("history2", "변동2");

            dataGridView1.RowHeadersWidth = 40;
            dataGridView1.Columns[0].Width = 60; //학번
            dataGridView1.Columns[1].Width = 70;
            dataGridView1.Columns[2].Width = 60;
            dataGridView1.Columns[3].Width = 60;
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[6].Width = 100;
            dataGridView1.Columns[7].Width = 70;
            dataGridView1.Columns[8].Width = 80;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 60;
            dataGridView1.Columns[11].Width = 60;

            dataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnHeadersHeight = 25;

            for (int i = 0; i <= 11; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            for (int i = 0; i <= 6; i++)
            {
                dataGridView1.Columns[i].ReadOnly = true;
                dataGridView1.Columns[i].DefaultCellStyle.SelectionForeColor = Color.Black;
            }
            dataGridView1.Columns[10].ReadOnly = true;
            dataGridView1.Columns[11].ReadOnly = true;
            dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.LightGray;

            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.LightGray;

            dataGridView1.Columns[1].DefaultCellStyle.SelectionBackColor = Color.AntiqueWhite;
            dataGridView1.Columns[2].DefaultCellStyle.SelectionBackColor = Color.AntiqueWhite;
            dataGridView1.Columns[3].DefaultCellStyle.SelectionBackColor = Color.LightGray;
            dataGridView1.Columns[4].DefaultCellStyle.SelectionBackColor = Color.LightGray;
            dataGridView1.Columns[5].DefaultCellStyle.SelectionBackColor = Color.LightGray;
            dataGridView1.Columns[6].DefaultCellStyle.SelectionBackColor = Color.LightGray;

            string connString = @"Data Source = " + dbFile + "; Pooling = true; FailIfMissing = false";
            SQLiteConnection conn = new SQLiteConnection(connString);

            using (conn)
            {
                conn.Open();
                string sql2 = "SELECT * FROM COMPANY ORDER BY CODE ASC";
                SQLiteCommand cmd2 = new SQLiteCommand(sql2, conn);
                SQLiteDataReader reader2 = cmd2.ExecuteReader();
                while (reader2.Read())
                {
                    string coCode = reader2.GetValue(0).ToString();
                    string coName = reader2.GetValue(1).ToString();
                    compList.Add(new Company { Code = coCode, Name = coName }); ;
                }
            }
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            int changeCnt = 0;
            int rowCnt = dataGridView1.Rows.Count;
            string strHead = txtCode1.Text.Trim();

            for (int i = 0; i < rowCnt; i++)
            {
                if (dataGridView1.Rows[i].HeaderCell.Value != null) changeCnt++;
            }

            DialogResult result;
            if (changeCnt > 0)
            {
                result = MessageBox.Show("저장하지 않은 자료가 있습니다. 저장하시겠습니까?", "자료 저장 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes) btnSave.PerformClick();
            }

            string connString = @"Data Source = " + dbFile + "; Pooling = true; FailIfMissing = false";
            SQLiteConnection conn = new SQLiteConnection(connString);

            using (conn)
            {
                conn.Open();
                dataGridView1.Rows.Clear();

                string sql = "SELECT * FROM STUDENT WHERE ID LIKE '" + strHead + "%' ORDER BY ID";
                SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string coCode, salary, jdate, udate, history1, history2;
                        string coName, coManager, coPhone;

                        string stID = reader.GetValue(0).ToString();
                        string stName = reader.GetValue(1).ToString();
                        string stSex = reader.GetValue(2).ToString();
                        string stSpecial = reader.GetValue(3).ToString();
                        if (stSpecial.Length > 0) stName = "*" + stName;

                        coName = ""; coManager = ""; coPhone = ""; salary = "";
                        coCode = ""; jdate = ""; udate = ""; history1 = ""; history2 = "";

                        string sql3 = "SELECT * FROM EMPLOY WHERE ID LIKE '" + stID + "%' ORDER BY ID ASC";
                        SQLiteCommand cmd3 = new SQLiteCommand(sql3, conn);
                        SQLiteDataReader reader3 = cmd3.ExecuteReader();

                        if (reader3.Read())
                        {
                            coCode = Convert.ToString(reader3.GetValue(1));
                            //coName = Convert.ToString(reader3.GetValue(2));
                            coName = findCompany(coCode);
                            salary = Convert.ToString(reader3.GetValue(3));
                            jdate = Convert.ToDateTime(reader3.GetValue(4)).ToString("yyyy-MM-dd");
                            udate = Convert.ToDateTime(reader3.GetValue(5)).ToString("yyyy-MM-dd");
                            history1 = reader3.GetValue(6).ToString();
                            history2 = reader3.GetValue(7).ToString();

                            string sql2 = "SELECT * FROM COMPANY WHERE CODE LIKE '" + coCode + "%' ORDER BY CODE";
                            SQLiteCommand cmd2 = new SQLiteCommand(sql2, conn);
                            SQLiteDataReader reader2 = cmd2.ExecuteReader();

                            if (reader2.Read())
                            {
                                //coName = reader2.GetValue(1).ToString();
                                coManager = reader2.GetValue(3).ToString();
                                coPhone = reader2.GetValue(2).ToString();
                            }
                            if (jdate == "1900-01-01") jdate = "";
                            if (udate == "1900-01-01") udate = "";
                        }

                        dataGridView1.Rows.Add(new object[] {
                            stID,  // 학번
                            stName,
                            stSex,
                            coCode,  // 기업체 코드
                            coName,
                            coManager,
                            coPhone,
                            salary,
                            jdate,
                            udate,
                            history1,
                            history2
                        });
                    }
                }
                conn.Close();
                dataGridView1.CurrentCell = null;
            }
        }

        private string findCompany(string coCode)
        {
            int index = -1;
            index = compList.FindIndex(x => x.Code == coCode);
            if (index < 0) return "";
            return compList[index].Name;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string fileName = "";
            SaveFileDialog saveFile = new SaveFileDialog();

            //saveFile.CreatePrompt = true;
            saveFile.OverwritePrompt = true;
            saveFile.InitialDirectory = @"C:\Data\";      // 최초 경로 설정
            saveFile.Title = "Excel 파일 저장";
            saveFile.DefaultExt = "xlsx";            // 기본 확장자
            saveFile.Filter = "엑셀파일(*.xlsx) | *.xlsx; *.xls;";

            if (saveFile.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("파일명을 입력하지 않았습니다.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            fileName = saveFile.FileName.ToString();

            Excel.Workbook excelBook = null;
            Excel.Worksheet excelSheet = null;
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null) return;

            excelApp.Visible = false;
            excelApp.DisplayAlerts = false; // 경고 메시지를 띄우지 않음
            excelApp.Interactive = false;   // 사용자의 조작에 방해받지 않음

            object misValue = System.Reflection.Missing.Value;

            try {
                excelBook = excelApp.Workbooks.Add(misValue);
                excelSheet = (Worksheet)excelBook.Worksheets.Add(misValue, misValue, misValue, misValue);

                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    excelSheet.Cells[1, j + 1] = dataGridView1.Columns[j].Name;
                }

                int row = 2;
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if (dataGridView1[3, i].Value.ToString() == "") continue;
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        excelSheet.Cells[row, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    row++;
                }

                excelApp.ActiveSheet.range("A1:L1").Interior.Color = Color.LightBlue;
                excelApp.ActiveSheet.Columns("A:L").AutoFit(); // 자동 열너비 
                Excel.Range usedRange = excelSheet.UsedRange;
                Excel.Range lastCell = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                Excel.Range totalRange = excelSheet.get_Range(excelSheet.get_Range("A1"), lastCell);
                totalRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                excelBook.SaveAs(fileName);
                excelBook.Close();
                excelApp.Quit();

                MessageBox.Show("파일을 성공적으로 저장하였습니다.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            } catch (Exception ex) {
                MessageBox.Show("\n파일 저장 취소 또는 오류 발생! : " + ex, "", MessageBoxButtons.OK, MessageBoxIcon.Error);

            } finally  {
                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);

                if (excelApp != null)
                {
                    Process[] pProcess;
                    pProcess = Process.GetProcessesByName("Excel");
                    pProcess[0].Kill();
                }
                excelSheet = null;
                excelBook = null;
                excelApp = null;
                GC.Collect();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
