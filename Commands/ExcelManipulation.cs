using CentralAptitudeTest.Models;
using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;   // 사용한 엑셀 객체들을 해제 해주기 위한 참조
using Microsoft.Office.Interop.Excel;   // 액셀 사용을 위한 참조
using System.ComponentModel;

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        // OpenFile() 에서 사용
        private Config Config;
        private BackgroundWorker Worker;

        private string outputAllSheetName = "부적응Data";

        private Application application;

        private Workbook InputDataWorkbook;
        private Workbook InputCollegeWorkbook;
        private Workbook OutputAllWorkbook;
        private Workbook OutputGraphWorkbook;

        private Worksheet InputDataWorksheet;
        private Worksheet InputCollegeWorksheet;
        private Worksheet OutputAllWorksheet;
        private Worksheet OutputGraphWorksheet;

        string DesktopPath;
        string Datetime = DateTime.Now.ToString("hhmmss");

        private Range CollegeListRange;
        private Range WholeInputDataRange;

        // ReadCollege() 에서 사용
        private List<string> CollegeList = new List<string>();
        private List<string> DepartList = new List<string>();

        private Dictionary<string, string> DepartCollegeDictionary = new Dictionary<string, string>();
        // 결과 작성 시 사용
        private Dictionary<string, int> ResultIndexDictionary = new Dictionary<string, int>();
        // 각 워크시트 옮겨 적을 때 사용
        private Dictionary<string, int> ResultRowcountDictionary = new Dictionary<string, int>();

        // 스트레스 취약성
        private int IndexStressColumn = 24;

        // PTSD
        private int IndexPtsdColumn = 15;

        // 편집증
        private int IndexParanoiaColumn = 19;

        // 정신증
        private int IndexPsychosisColumn = 20;

        // 우울
        private int IndexDepressedColumn = 7;

        // 불안
        private int IndexUnrestColumn = 8;

        // 중독
        private int IndexAddictionColumn = 22;

        // 공포불안
        private int IndexFearColumn = 9;

        // 분노공격
        private int IndexAngerColumn = 16;

        // 조증
        private int IndexManiaColumn = 18;

        List<int> PreventAptitudeRecList = new List<int>();
        List<int> PreventStressList = new List<int>();
        List<int> PreventTraumaList = new List<int>();
        List<int> PreventIsolateList = new List<int>();
        List<int> PreventIPConflictList = new List<int>();

        List<int> SeriousAptitudeRecList = new List<int>();
        List<int> SeriousStressList = new List<int>();
        List<int> SeriousTraumaList = new List<int>();
        List<int> SeriousIsolateList = new List<int>();
        List<int> SeriousIPConflictList = new List<int>();

        public ExcelManipulation(Config config, BackgroundWorker worker)
        {
            Worker = worker;
            worker.ReportProgress(10, String.Format("생성자 동작 시작"));

            // 바탕화면 경로 불러오기
            DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Excel 파일 저장 경로 및 파일 이름 설정
            // string 내부에서  {0}을 통한 combine 실패. 해당 코드 사용 불가!!!
            // path = Path.Combine(DesktopPath, "{0}.xlsx", Datetime);

            Config = config;
            application = new Application();

            OpenFile(config);
            worker.ReportProgress(11, String.Format("생성자 동작 종료"));
        }


        public void OpenFile(Config config)
        {
            Worker.ReportProgress(12, String.Format("파일 열기 시작"));
            // 입력 Excel 파일(워크북) 불러오기
            InputDataWorkbook = application.Workbooks.Open(config.FilePath.whole_data_filePath);
            InputCollegeWorkbook = application.Workbooks.Open(config.FilePath.process_data_filePath);

            Debug.WriteLine("입력 데이터 활성화 worksheet : " + InputDataWorkbook.Worksheets.Count);

            // Excel 화면 창 띄우기
            // application.Visible = true;

            OutputAllWorkbook = application.Workbooks.Add();
            OutputGraphWorkbook = application.Workbooks.Add();      

            // worksheet 생성하기
            InputDataWorksheet = (Worksheet)InputDataWorkbook.Sheets[1];
            InputCollegeWorksheet = (Worksheet)InputCollegeWorkbook.Sheets[1];
            OutputAllWorksheet = (Worksheet)OutputAllWorkbook.ActiveSheet;
            OutputGraphWorksheet = (Worksheet)OutputGraphWorkbook.ActiveSheet;

            // 전체 입력 data 영역 설정
            WholeInputDataRange = InputDataWorksheet.UsedRange;

            // 단과 대학 영역 설정
            CollegeListRange = InputCollegeWorksheet.UsedRange;

            Worker.ReportProgress(13, String.Format("파일 열기 종료"));
        }

        public void CloseFile()
        {
            Worker.ReportProgress(97, String.Format("작업 완료, 파일 닫기 시작!!!"));

            Debug.WriteLine("=============================작업 완료, 파일 닫기 시작=============================");

            // Save -> Close 순으로 수행
            string allfilenaming = "전체" + Datetime + ".xlsx";
            string graphfilenaming = "그래프" + Datetime + ".xlsx";

            string allpath = Path.Combine(DesktopPath, allfilenaming);
            string graphpath = Path.Combine(DesktopPath, graphfilenaming);

            // Save -> SaveAs 순으로 수행
            OutputAllWorkbook.SaveAs(Filename: allpath);
            OutputGraphWorkbook.SaveAs(Filename: graphpath);

            InputDataWorkbook.Close();
            InputCollegeWorkbook.Close();
            OutputAllWorkbook.Close();
            OutputGraphWorkbook.Close();

            application.Quit();

            // background에서 실행중인 객체들 마저 확실하게 해제시켜주기 위하여 사용.
            Marshal.ReleaseComObject(InputDataWorksheet);
            Marshal.ReleaseComObject(InputCollegeWorksheet);
            Marshal.ReleaseComObject(OutputAllWorksheet);
            Marshal.ReleaseComObject(OutputGraphWorksheet);

            Marshal.ReleaseComObject(InputDataWorkbook);
            Marshal.ReleaseComObject(InputCollegeWorkbook);
            Marshal.ReleaseComObject(OutputAllWorkbook);
            Marshal.ReleaseComObject(OutputGraphWorkbook);

            Marshal.ReleaseComObject(application);

            GC.Collect();

            Worker.ReportProgress(100, String.Format("작업 파일 종료중..."));

            Debug.WriteLine("=============================작업 완료, 파일 닫기 종료=============================");
        }

        // 1번째 수행 함수
        public void ReadCollege()
        {
            Debug.WriteLine("=============================단과대학 및 학과 읽기 시작=============================");
            Worker.ReportProgress(14, String.Format("단과대학 및 학과 읽기 시작"));

            var depart = "";
            var college = "";
            var collegeRow = CollegeListRange.Rows.Count;
            var collegeColumn = CollegeListRange.Columns.Count;
            var index = 2;

            // 첫째줄 제목을 지우기 위해 row=2부터 시작
            Debug.WriteLine("=============================ClassData 딕셔너리 생성 시작=============================");
            for (int row = 2; row <= collegeRow; row++)
            {
                depart = (string)(CollegeListRange.Cells[row, 2] as Range).Value2;
                college = (string)(CollegeListRange.Cells[row, 1] as Range).Value2;

                if (college != null)
                {
                    college = (string)(CollegeListRange.Cells[row, 1] as Range).Value2;
                    Debug.WriteLine(college + " 처음찾기 성공");

                    Debug.WriteLine("기입될 데이터 : " + depart + "---" + college);
                    CollegeList.Add(college);

                    Debug.WriteLine(depart + " 딕셔너리 작성");

                    index = row;
                }
                else
                {
                    Debug.WriteLine(college + "---중복");

                    college = (string)(CollegeListRange.Cells[index, 1] as Range).Value2;

                    Debug.WriteLine("기입될 데이터 : " + depart + "---" + college);
                    Debug.WriteLine(depart + " 딕셔너리 작성");
                }
                DepartCollegeDictionary.Add(depart, college);

                // Excel에 값 삽입하는 기본 문법
                // Range rg1 = (Range)OutputAllWorksheet.Cells[1, 1];
                // rg1.Value = "hello world";       
            }
            Debug.WriteLine("=============================ClassData 딕셔너리 생성 완료=============================");
            Worker.ReportProgress(17, String.Format("단과대학 및 학과 읽는 중."));

            // CollegeList 에서 null값 전부 제거
            CollegeList.RemoveAll(item => item == null);

            // 결과 엑셀에 학과별 worksheet 생성
            for (int workSheetNum = 0; workSheetNum < CollegeList.Count; workSheetNum++)
            {
                OutputAllWorkbook.Worksheets.Add(After: OutputAllWorkbook.Worksheets[workSheetNum + 1]);
                var currentWorksheet = OutputAllWorkbook.Worksheets.Item[workSheetNum + 1] as Worksheet;

                currentWorksheet.Name = CollegeList[workSheetNum];
                ResultRowcountDictionary.Add(CollegeList[workSheetNum], 1);

                Debug.WriteLine(CollegeList[workSheetNum] + "로 워크 시트 이름 변경 성공!");

                Worker.ReportProgress(23, String.Format("단과대학 및 학과 읽는 중..."));
            }

            var lastWorksheet = OutputAllWorkbook.Worksheets.Item[OutputAllWorkbook.Worksheets.Count] as Worksheet;
            lastWorksheet.Name = outputAllSheetName;

            Worker.ReportProgress(25, String.Format("단과대학 및 학과 읽기 종료"));
            Debug.WriteLine("=============================단과대학 및 학과 읽기 종료=============================");
        }

        // 2번째 수행 함수
        public void MisfitFiltering()
        {
            Worker.ReportProgress(26, String.Format("부적응자 필터링 시작"));

            Debug.WriteLine("=============================부적응자 필터링 시작=============================");
            var preventAptitudeRecList = new List<int>();
            var preventStressList = new List<int>();
            var preventTraumaList = new List<int>();
            var preventIsolateList = new List<int>();
            var preventIPConflictList = new List<int>();

            var seriousAptitudeRecList = new List<int>();
            var seriousStressList = new List<int>();
            var seriousTraumaList = new List<int>();
            var seriousIsolateList = new List<int>();
            var seriousIPConflictList = new List<int>();

            // 전체 데이터 처음 인자부터 돌면서 문제되는 열 탐색.
            for (int rowCount = 2; rowCount < WholeInputDataRange.Rows.Count; rowCount++)
            {
                var targetValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexStressColumn] as Range).Value2);

                var paranoiaValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexParanoiaColumn] as Range).Value2);
                var psychosisValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexPsychosisColumn] as Range).Value2);
                var depressedValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexDepressedColumn] as Range).Value2);
                var unrestValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexUnrestColumn] as Range).Value2);
                var ptsdValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexPtsdColumn] as Range).Value2);
                var addictionValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexAddictionColumn] as Range).Value2);
                var fearValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexFearColumn] as Range).Value2);
                var maniaValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexManiaColumn] as Range).Value2);
                var angerValue = Convert.ToInt32((WholeInputDataRange.Cells[rowCount, IndexAngerColumn] as Range).Value2);

                // 예방집단
                if (targetValue >= 60 && targetValue < 70)
                {
                    // 적성인식
                    if (paranoiaValue >= 60 || psychosisValue >= 60)
                    {
                        Debug.WriteLine("적성인식-예방 열정보 삽입");
                        preventAptitudeRecList.Add(rowCount);
                    }

                    // 스트레스
                    if (depressedValue >= 60 || unrestValue >= 60)
                    {
                        Debug.WriteLine("스트레스-예방 열정보 삽입");
                        preventStressList.Add(rowCount);
                    }

                    // 외상경험
                    if (ptsdValue >= 60 || addictionValue >= 60)
                    {
                        Debug.WriteLine("외상경험-예방 열정보 삽입");
                        preventTraumaList.Add(rowCount);
                    }

                    // 고립
                    if (fearValue >= 60 || paranoiaValue >= 60)
                    {
                        Debug.WriteLine("고립-예방 열정보 삽입");
                        preventIsolateList.Add(rowCount);
                    }

                    // 대인갈등
                    if (maniaValue >= 60 || angerValue >= 60)
                    {
                        Debug.WriteLine("대인갈등-예방 열정보 삽입");
                        preventIPConflictList.Add(rowCount);
                    }
                }

                // 문제집단
                if (targetValue >= 70)
                {
                    // 적성인식
                    if (paranoiaValue >= 70 || psychosisValue >= 70)
                    {
                        Debug.WriteLine("적성인식-문제 열정보 삽입");
                        seriousAptitudeRecList.Add(rowCount);
                    }

                    // 스트레스
                    if (depressedValue >= 70 || unrestValue >= 70)
                    {
                        Debug.WriteLine("스트레스-문제 열정보 삽입");
                        seriousStressList.Add(rowCount);
                    }

                    // 외상경험
                    if (ptsdValue >= 70 || addictionValue >= 70)
                    {
                        Debug.WriteLine("외상경험-문제 열정보 삽입");
                        seriousTraumaList.Add(rowCount);
                    }

                    // 고립
                    if (fearValue >= 70 || paranoiaValue >= 70)
                    {
                        Debug.WriteLine("고립-문제 열정보 삽입");
                        seriousIsolateList.Add(rowCount);
                    }

                    // 대인갈등
                    if (maniaValue >= 70 || angerValue >= 70)
                    {
                        Debug.WriteLine("대인갈등-문제 열정보 삽입");
                        seriousIPConflictList.Add(rowCount);
                    }
                }
                if (rowCount == rowCount / 3)
                {
                    Worker.ReportProgress(28, String.Format("부적응자 필터링 중."));
                }

                if (rowCount == rowCount / 2)
                {
                    Worker.ReportProgress(30, String.Format("부적응자 필터링 중.."));
                }

                if (rowCount == (rowCount / 3) * 2)
                {
                    Worker.ReportProgress(32, String.Format("부적응자 필터링 중..."));
                }
            }
            PreventAptitudeRecList = preventAptitudeRecList;
            PreventStressList = preventStressList;
            PreventTraumaList = preventTraumaList;
            PreventIsolateList = preventIsolateList;
            PreventIPConflictList = preventIPConflictList;

            SeriousAptitudeRecList = seriousAptitudeRecList;
            SeriousStressList = seriousStressList;
            SeriousTraumaList = seriousTraumaList;
            SeriousIsolateList = seriousIsolateList;
            SeriousIPConflictList = seriousIPConflictList;

            Debug.WriteLine("=============================부적응자 필터링 종료=============================");

            MisfitPreventWriting();
        }

        // 3번째 수행 함수
        public string SeparateEachDepart()
        {
            Worker.ReportProgress(36, String.Format("단과대학 별 학과 분류 데이터 기입 시작"));
            Debug.WriteLine("=============================단과대별 학과 분류하여 워크시트 데이터 기입 시작=============================");
                       
            var inputdataRowCount = InputDataWorksheet.Rows.Count;

            ResultRowcountDictionary.Add("전체Data", WholeInputDataRange.Rows.Count);

            InputDataWorksheet.Copy(OutputAllWorkbook.Worksheets[OutputAllWorkbook.Worksheets.Count]);

            for (var index = 2; index <= inputdataRowCount; index++)
            {
                var currentdepartname = (string)(InputDataWorksheet.Cells[index, 1] as Range).Value2;

                if(currentdepartname == null)
                {
                    break;
                }

                Debug.WriteLine("currentdepartname : " + currentdepartname);

                if(DepartCollegeDictionary.ContainsKey(currentdepartname))
                {
                    var currentcollegename = DepartCollegeDictionary[currentdepartname];
                    Debug.WriteLine("currentcollegename : " + currentcollegename);
                    var writeworksheet = OutputAllWorkbook.Worksheets.Item[currentcollegename] as Worksheet;
                    var writerowindex = ResultRowcountDictionary[currentcollegename];
                    ResultRowcountDictionary[currentcollegename] += 1;
                    var toindex = "A" + writerowindex + ":Z" + writerowindex;
                    var fromindex = "A" + index + ":Z" + index;

                    Debug.WriteLine("toindex : " + toindex);
                    Debug.WriteLine("fromindex : " + fromindex);

                    var from = InputDataWorksheet.UsedRange.Range[fromindex];
                    var to = writeworksheet.Range[toindex];
                    from.Copy(to);
                    Debug.WriteLine("복사 성공...!");
                }
                else
                {
                    Debug.WriteLine("없는 학과 명 입니다." + currentdepartname);
                    Worker.ReportProgress(0,String.Format("없는 학과명 발생!!!!!! 입력 데이터를 확인해 주세요!"));
                    return currentdepartname;
                }
            }

            for(var index = 0; index < ResultRowcountDictionary.Keys.Count; index++)
            {
                var keylist = ResultRowcountDictionary.Keys.ToList();
                var key = keylist[index];
                ResultRowcountDictionary[key] -= 1;
            }

            Debug.WriteLine("=============================단과대별 학과 분류하여 워크시트 데이터 기입 종료=============================");
            Worker.ReportProgress(68, String.Format("단과대학 별 학과 분류 데이터 기입 종료"));

            return null;
        }

        // 4번째 수행 함수
        // StudentNum 딕셔너리 생성이 아직 안되었기에, SeparateEachDepart 다음에 호출
        public void GraphFileTask()
        {
            Worker.ReportProgress(69, String.Format("그래프 파일 작업 시작"));
            Debug.WriteLine("=============================그래프 파일 시작=============================");

            var graphSheet = OutputGraphWorkbook.Worksheets.Item[1] as Worksheet;
            graphSheet.Name = "그래프Data";

            var from = WholeInputDataRange.Range["E1:Z1"];
            var to = graphSheet.Range["B1:W1"];

            from.Copy(to);

            CollegeList.Reverse();
            CollegeList.Add("전체Data");
            CollegeList.Reverse();

            for (var index = 0; index < CollegeList.Count; index++)
            {
                var college = CollegeList[index];
                if (ResultRowcountDictionary[college] < 2)
                {
                    CollegeList.Remove(college);
                    var worksheet = OutputAllWorkbook.Worksheets.Item[college] as Worksheet;
                    worksheet.Delete();
                }
            }

            var collegelist = CollegeList;

            var studentkeysindex = 0;

            for (int graphcollegeindex = 2; graphcollegeindex < CollegeList.Count * 2 + 1; graphcollegeindex += 2)
            {
                var inputtitle = collegelist[studentkeysindex] + "(n=" + ResultRowcountDictionary[collegelist[studentkeysindex]] + ")";

                if (ResultRowcountDictionary[collegelist[studentkeysindex]] >= 2)
                {
                    var targetCell = (graphSheet.Cells[graphcollegeindex, 1] as Range);
                    targetCell.Value = inputtitle;

                    targetCell = (graphSheet.Cells[graphcollegeindex + 1, 1] as Range);
                    targetCell.Value = inputtitle;
                }

                studentkeysindex++;
            }
            Worker.ReportProgress(70, String.Format("그래프 파일 작업 종료"));

            Debug.WriteLine("=============================그래프 파일 종료=============================");
        }

        // 5번째 수행 함수
        public void ResultEachCollege()
        {
            Worker.ReportProgress(71, String.Format("최종 결과 작성 작업 시작"));
            Debug.WriteLine("===각 대학 결과 정보 기입===");

            var collegelist = CollegeList;

            var progressCount = 1;

            foreach(var college in collegelist)
            {
                Debug.WriteLine(college + " 결과 작성 시작");

                if(college == "전체Data")
                {
                    continue;
                }

                if(progressCount == collegelist.Count / 3)
                {
                    Worker.ReportProgress(76, String.Format("최종 결과 작성 작업 중."));
                }
                if (progressCount == collegelist.Count / 2)
                {
                    Worker.ReportProgress(81, String.Format("최종 결과 작성 작업 중.."));
                }
                if (progressCount == (collegelist.Count / 3) * 2)
                {
                    Worker.ReportProgress(86, String.Format("최종 결과 작성 작업 중..."));
                }

                var targetworksheet = OutputAllWorkbook.Worksheets.Item[college] as Worksheet;

                // 총 단과대 인원 작성
                var studentcountindex = targetworksheet.UsedRange.Rows.Count + 3;

                ResultIndexDictionary.Add(college, studentcountindex);

                var writeplace = targetworksheet.Range[targetworksheet.Cells[studentcountindex, 5], targetworksheet.Cells[studentcountindex, 26]];
                
                writeplace.Value2 = ResultRowcountDictionary[college];

                // 개별 단과대 개별 이상자 인원수 파악
                studentcountindex -= 1; // 카운트 낱개 갯수 위치 조정

                // 마지막 끝나는 studentcountindex == 컬럼 갯수 위치
                for (var columnindex = 5;  columnindex < 27; columnindex++)
                {
                    var columncount = ColumnCounter(targetworksheet, columnindex, college);
                    (targetworksheet.Cells[studentcountindex, columnindex] as Range).Value2 = columncount;

                    studentcountindex += 2;

                    var input = Math.Round((float)columncount / ResultRowcountDictionary[college], 4);
                    (targetworksheet.Cells[studentcountindex, columnindex] as Range).Value2 = input;
                    
                    studentcountindex -= 2;
                }
            }

            GraphResultWriting();
        }

        private void GraphResultWriting()
        {
            var startIndex = 4;
            var graphworksheet = OutputGraphWorkbook.Worksheets.Item["그래프Data"] as Worksheet;
            var contentsList = ResultIndexDictionary.Keys;

            foreach(var contents in contentsList)
            {
                var originalIndex = ResultIndexDictionary[contents] - 1;
                var targetworksheet = OutputAllWorkbook.Worksheets.Item[contents] as Worksheet;

                var fromIndex = "E" + originalIndex + ":Z" + originalIndex;
                var toIndex = "B" + startIndex + ":W" + startIndex;

                var from = targetworksheet.Range[fromIndex];
                var to = graphworksheet.Range[toIndex];
                from.Copy(to);

                // 퍼센테이지 copy

                startIndex++;
                originalIndex += 2;

                fromIndex = "E" + originalIndex + ":Z" + originalIndex;
                toIndex = "B" + startIndex + ":W" + startIndex;

                from = targetworksheet.Range[fromIndex];
                to = graphworksheet.Range[toIndex];
                from.Copy(to);

                startIndex++;
            }

            for(var columnIndex = 2; columnIndex < graphworksheet.UsedRange.Columns.Count+1; columnIndex++)
            {
                var inputValue = 0;

                for(var rowIndex = 4; rowIndex < graphworksheet.UsedRange.Rows.Count; rowIndex+=2)
                {
                    var temp = Convert.ToInt32((graphworksheet.Cells[rowIndex, columnIndex] as Range).Value2);
                    inputValue += temp;
                }

                (graphworksheet.Cells[2, columnIndex] as Range).Value2 = inputValue;

                var inputPercentageValue = Math.Round((float)inputValue / ResultRowcountDictionary["전체Data"], 4);

                (graphworksheet.Cells[3, columnIndex] as Range).Value2 = inputPercentageValue;
            }

            Worker.ReportProgress(90, String.Format("최종 결과 작성 작업 시작"));
        }

        private int ColumnCounter(Worksheet targetworksheet, int columnindex, string college)
        {
            var count = 0;

            if(college == "전체Data")
            {
                for (var index = 2; index < ResultRowcountDictionary[college]; index++)
                {
                    Range range = targetworksheet.Cells[index, columnindex] as Range;
                    var temp = Convert.ToInt32(range.Value2);

                    if (temp >= 70)
                    {
                        count++;
                    }
                }
            }
            else
            {
                for (var index = 1; index < ResultRowcountDictionary[college]; index++)
                {
                    Range range = targetworksheet.Cells[index, columnindex] as Range;
                    var temp = Convert.ToInt32(range.Value2);

                    if (temp >= 70)
                    {
                        count++;
                    }
                }
            }            

            return count;
        }        

        private void MisfitPreventWriting()
        {
            Debug.WriteLine("=============================부적응자 데이터 작성 시작=============================");

            int rowIndex = 1;
            rowIndex = MisfitWrite(false, rowIndex, IndexParanoiaColumn, IndexPsychosisColumn, "적성인식-예방");
            rowIndex = MisfitWrite(false, rowIndex, IndexDepressedColumn, IndexUnrestColumn, "스트레스-예방");
            rowIndex = MisfitWrite(false, rowIndex, IndexPtsdColumn, IndexAddictionColumn, "외상경험-예방");
            rowIndex = MisfitWrite(false, rowIndex, IndexFearColumn, IndexParanoiaColumn, "고립-예방");
            rowIndex = MisfitWrite(false, rowIndex, IndexManiaColumn, IndexAngerColumn, "대인갈등-예방");

            rowIndex = 1;
            rowIndex = MisfitWrite(true, rowIndex, IndexParanoiaColumn, IndexPsychosisColumn, "적성인식-문제");
            rowIndex = MisfitWrite(true, rowIndex, IndexDepressedColumn, IndexUnrestColumn, "스트레스-문제");
            rowIndex = MisfitWrite(true, rowIndex, IndexPtsdColumn, IndexAddictionColumn, "외상경험-문제");
            rowIndex = MisfitWrite(true, rowIndex, IndexFearColumn, IndexParanoiaColumn, "고립-문제");
            rowIndex = MisfitWrite(true, rowIndex, IndexManiaColumn, IndexAngerColumn, "대인갈등-문제");

            Debug.WriteLine("=============================부적응자 데이터 작성 완료=============================");
            Worker.ReportProgress(35, String.Format("부적응자 필터링 종료"));
        }

        private int MisfitWrite(Boolean isSerious, int rowIndex, int target1, int target2, string title)
        {
            var targetWorksheet = OutputAllWorkbook.Worksheets.Item[outputAllSheetName] as Worksheet;

            var currentRowIndex = rowIndex;
            var preventTitleColumnIndex = 1;

            var targetlist = new List<int>();

            var departIndex = 1;
            var numIndex = 2;
            var nameIndex = 3;
            var sexIndex = 4;
            var target1Index = 5;
            var target2Index = 6;
            var stressIndex = 7;

            if (isSerious)
            {
                preventTitleColumnIndex = 10;
                departIndex = 10;
                numIndex = 11;
                nameIndex = 12;
                sexIndex = 13;
                target1Index = 14;
                target2Index = 15;
                stressIndex = 16;

                if (title.Contains("적성인식"))
                {
                    targetlist = SeriousAptitudeRecList;
                }
                if (title.Contains("스트레스"))
                {
                    targetlist = SeriousStressList;
                }
                if (title.Contains("외상경험"))
                {
                    targetlist = SeriousTraumaList;
                }
                if (title.Contains("고립"))
                {
                    targetlist = SeriousIsolateList;
                }
                if (title.Contains("대인갈등"))
                {
                    targetlist = SeriousIPConflictList;
                }
            }
            else
            {
                if (title.Contains("적성인식"))
                {
                    targetlist = PreventAptitudeRecList;
                }
                if (title.Contains("스트레스"))
                {
                    targetlist = PreventStressList;
                }
                if (title.Contains("외상경험"))
                {
                    targetlist = PreventTraumaList;
                }
                if (title.Contains("고립"))
                {
                    targetlist = PreventIsolateList;
                }
                if (title.Contains("대인갈등"))
                {
                    targetlist = PreventIPConflictList;
                }
            }

            (targetWorksheet.Cells[currentRowIndex, preventTitleColumnIndex] as Range).Value = title;

            currentRowIndex += 2;
            // 각주 쓰기
            (targetWorksheet.Cells[currentRowIndex, departIndex] as Range).Value = "학과";
            (targetWorksheet.Cells[currentRowIndex, numIndex] as Range).Value = "학번";
            (targetWorksheet.Cells[currentRowIndex, nameIndex] as Range).Value = "성명";
            (targetWorksheet.Cells[currentRowIndex, sexIndex] as Range).Value = "성별";
            (targetWorksheet.Cells[currentRowIndex, target1Index] as Range).Value = (WholeInputDataRange.Cells[1, target1] as Range).Value2;
            (targetWorksheet.Cells[currentRowIndex, target2Index] as Range).Value = (WholeInputDataRange.Cells[1, target2] as Range).Value2;
            (targetWorksheet.Cells[currentRowIndex, stressIndex] as Range).Value = "스트레스취약성";

            currentRowIndex++;

            foreach (var index in targetlist)
            {
                (targetWorksheet.Cells[currentRowIndex, departIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 1] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, numIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 2] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, nameIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 3] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, sexIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 4] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, target1Index] as Range).Value2 = (WholeInputDataRange.Cells[index, target1] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, target2Index] as Range).Value2 = (WholeInputDataRange.Cells[index, target2] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, stressIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 24] as Range).Value2;

                currentRowIndex++;
            }

            currentRowIndex += 2;

            return currentRowIndex;
        }
    }
}