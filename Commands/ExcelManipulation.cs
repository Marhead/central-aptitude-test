using CentralAptitudeTest.Models;
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;   // 사용한 엑셀 객체들을 해제 해주기 위한 참조
using Microsoft.Office.Interop.Excel;   // 액셀 사용을 위한 참조

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        // OpenFile() 에서 사용
        private Config Config;

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
        private Dictionary<string, List<string>> ClassData = new Dictionary<string, List<string>>();
        private Dictionary<string, int> StudentNum = new Dictionary<string, int>();

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

        public ExcelManipulation(Config config)
        {
            Debug.WriteLine("=============================생성자 동작 시작=============================");

            // 바탕화면 경로 불러오기
            DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Excel 파일 저장 경로 및 파일 이름 설정
            // string 내부에서  {0}을 통한 combine 실패. 해당 코드 사용 불가!!!
            // path = Path.Combine(DesktopPath, "{0}.xlsx", Datetime);

            Config = config;
            application = new Application();

            OpenFile(config);
            Debug.WriteLine("=============================생성자 동작 종료=============================");
        }


        public void OpenFile(Config config)
        {
            Debug.WriteLine("=============================파일 열기 시작=============================");
            // 입력 Excel 파일(워크북) 불러오기
            InputDataWorkbook = application.Workbooks.Open(config.FilePath.whole_data_filePath);
            InputCollegeWorkbook = application.Workbooks.Open(config.FilePath.process_data_filePath);

            Debug.WriteLine("입력 데이터 활성화 worksheet : " + InputDataWorkbook.Worksheets.Count);

            // Excel 화면 창 띄우기
            // application.Visible = true;

            OutputAllWorkbook = application.Workbooks.Add();
            OutputGraphWorkbook = application.Workbooks.Add();
               
            // Test용
            // 기존 Excel 파일(워크북) 불러오기
            //OutputAllWorkbook = application.Workbooks.Open(@"C:\\code\\대구가톨릭대학교전체정리.xlsx");
            //OutputGraphWorkbook = application.Workbooks.Open(@"C:\\code\\대구가톨릭대학그래프정리.xlsx");         

            // worksheet 생성하기
            InputDataWorksheet = (Worksheet)InputDataWorkbook.Sheets[1];
            InputCollegeWorksheet = (Worksheet)InputCollegeWorkbook.Sheets[1];
            OutputAllWorksheet = (Worksheet)OutputAllWorkbook.ActiveSheet;
            OutputGraphWorksheet = (Worksheet)OutputGraphWorkbook.ActiveSheet;

            // 전체 입력 data 영역 설정
            WholeInputDataRange = InputDataWorksheet.UsedRange;

            // 단과 대학 영역 설정
            CollegeListRange = InputCollegeWorksheet.UsedRange;
            
            Debug.WriteLine("=============================파일 열기 종료=============================");
        }

        public void CloseFile()
        {
            Debug.WriteLine("=============================작업 완료, 파일 닫기 시작=============================");

            // Save -> Close 순으로 수행
            string allfilenaming = "전체" + Datetime + ".xlsx";
            string graphfilenaming = "그래프" + Datetime + ".xlsx";

            string allpath = Path.Combine(DesktopPath, allfilenaming);
            string graphpath = Path.Combine(DesktopPath, graphfilenaming);

            // Save -> SaveAs 순으로 수행
            // InputDataWorkbook.Save();
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
        }

        // summary
        // 2번째 입력파일에서 부터 각 "단대"와 "학과"를 읽어오기
        // 읽어온 데이터로, 전체 데이터 "워크시트" 생성하기
        public void ReadCollege()
        {
            Debug.WriteLine("=============================단과대학 및 학과 읽기 시작=============================");

            var Depart = "";
            var College = "";
            var DepartStartIndexList = new List<int>();
            var CollegeRow = CollegeListRange.Rows.Count;
            var CollegeColumn = CollegeListRange.Columns.Count;

            // 첫째줄 제목을 지우기 위해 row=2부터 시작
            for(int row = 2; row <= CollegeRow; row++)
            {
                Depart = (string)(CollegeListRange.Cells[row, 2] as Range).Value2;
                College = (string)(CollegeListRange.Cells[row, 1] as Range).Value2;

                CollegeList.Add(College);
                if(College != null)
                {
                    Debug.WriteLine("단과대 : " + College);
                }
                // Range collegeinput = (Range)OutputAllWorksheet.Cells[row, 1];
                // collegeinput.Value = (string)(CollegeListRange.Cells[row, 1] as Range).Value2;

                if (College != null)
                {
                    DepartStartIndexList.Add(row-2);
                    Debug.WriteLine("입력 row 수 : {0}", row);
                }

                DepartList.Add(Depart);
                Debug.WriteLine("학과 : " + Depart);
                // Range departinput = (Range)OutputAllWorksheet.Cells[row, 2];
                // departinput.Value = (string)(CollegeListRange.Cells[row, 2] as Range).Value2;                
            }

            // CollegeList 에서 null값 전부 제거
            CollegeList.RemoveAll(item => item == null);

            DepartStartIndexList.Add(CollegeRow-1);

            // 딕셔너리 생성 loop
            Debug.WriteLine("=============================ClassData 딕셔너리 생성 시작=============================");
            for (int DepartIndex = 0; DepartIndex < DepartStartIndexList.Count-1; DepartIndex++)
            {
                // 임시 리스트 초기화
                var DictInputDepartList = new List<string>();

                // 딕셔너리 Value 생성
                // GetRange( int 시작인덱스, int 갯수 )
                DictInputDepartList = DepartList.GetRange(DepartStartIndexList[DepartIndex], DepartStartIndexList[DepartIndex + 1] - DepartStartIndexList[DepartIndex]);

                Debug.WriteLine("주입할 대학 이름 : " + CollegeList[DepartIndex]);
                DictInputDepartList.ForEach(item => Debug.WriteLine("주입할 학부 이름 : " + item));

                Debug.WriteLine("***Dictionary 데이터 주입 준비***");

                ClassData.Add(CollegeList[DepartIndex], DictInputDepartList);

                Debug.WriteLine("***Dictionary 데이터 주입 완료***");
            }

            Debug.WriteLine("=============================ClassData 딕셔너리 생성 완료=============================");

            // ClassData Dictionary 검사 부분
            foreach(KeyValuePair<string, List<string>> items in ClassData)
            {
                Debug.WriteLine(items.Key);
                ClassData[items.Key].ForEach(depart => Debug.WriteLine(depart));
            }

            // 결과 엑셀에 학과별 worksheet 생성
            for (int workSheetNum = 0; workSheetNum < ClassData.Keys.Count; workSheetNum++)
            {
                OutputAllWorkbook.Worksheets.Add(After: OutputAllWorkbook.Worksheets[workSheetNum + 1]);
                var currentWorksheet = OutputAllWorkbook.Worksheets.Item[workSheetNum + 1] as Worksheet;

                currentWorksheet.Name = CollegeList[workSheetNum];

                Debug.WriteLine(CollegeList[workSheetNum] + "로 워크 시트 이름 변경 성공!");
            }

            var lastWorksheet = OutputAllWorkbook.Worksheets.Item[OutputAllWorkbook.Worksheets.Count] as Worksheet;
            lastWorksheet.Name = outputAllSheetName;

            Debug.WriteLine("=============================단과대학 및 학과 읽기 종료=============================");
        }

        // StudentNum 딕셔너리 생성이 아직 안되었기에, SeparateEachDepart 다음에 호출
        public void GraphFileTask()
        {
            Debug.WriteLine("=============================그래프 파일 시작=============================");

            var graphSheet = OutputGraphWorkbook.Worksheets.Item[1] as Worksheet;
            graphSheet.Name = "그래프Data";

            var from = WholeInputDataRange.Range["E1:Z1"];
            var to = graphSheet.Range["B1:W1"];

            from.Copy(to);

            var graphinputcollegelist = new List<string>(StudentNum.Keys);

            var studentkeysindex = 0;

            for(int graphcollegeindex = 2; graphcollegeindex < StudentNum.Keys.Count * 2 + 1; graphcollegeindex+=2)
            {
                var inputtitle = graphinputcollegelist[studentkeysindex] + "(n=" + StudentNum[graphinputcollegelist[studentkeysindex]] + ")";
                
                if(StudentNum[graphinputcollegelist[studentkeysindex]] >= 2)
                {
                    var targetCell = (graphSheet.Cells[graphcollegeindex, 1] as Range);
                    targetCell.Value = inputtitle;

                    targetCell = (graphSheet.Cells[graphcollegeindex + 1, 1] as Range);
                    targetCell.Value = inputtitle;
                }

                studentkeysindex++;
            }

            Debug.WriteLine("=============================그래프 파일 종료=============================");
        }

        public void SeparateEachDepart()
        {
            Debug.WriteLine("=============================단과대별 학과 분류하여 워크시트 데이터 기입 시작=============================");
            
            // Excel에 값 삽입하는 기본 문법
            // Range rg1 = (Range)OutputAllWorksheet.Cells[1, 1];
            // rg1.Value = "hello world";

            var dataRowNum = WholeInputDataRange.Rows.Count;

            // 단과대 모음 리스트
            var collegeNameList = ClassData.Keys;
            StudentNum.Add("전체인원", WholeInputDataRange.Rows.Count);

            foreach (string collegename in collegeNameList)
            {
                // 출력 워크시트 설정
                var targetWorksheet = OutputAllWorkbook.Worksheets.Item[collegename] as Worksheet;

                Debug.WriteLine(collegename + " 학부 기재 작업 준비");

                // 단과대 학부 리스트 불러오기
                var departNameList = ClassData[collegename];

                var isFirst = true;

                foreach (string departName in departNameList)
                {
                    Debug.WriteLine(departName + " 작업 시작!!!");

                    // 엑셀 데이터 쓰기 위치 값
                    int nextStartIndex = targetWorksheet.UsedRange.Rows.Count + 1;

                    if (isFirst)
                    {
                        nextStartIndex -= 1;
                        isFirst = false;
                    }

                    Debug.WriteLine(departName + " 기입 시작 위치 : {0}", nextStartIndex);

                    // 복사 시작 인덱스 위치 판별
                    var isCopyStartIndex = true;

                    // 복사 시작 인덱스 위치값
                    int copyStartIndex = 0;

                    // 복사 종료 인덱스 위치값
                    int copyEndIndex = 0;

                    // from 위치값 찾기
                    for (int rowCount = 1; rowCount < dataRowNum; rowCount++)
                    {
                        if (departName == (string)(WholeInputDataRange.Cells[rowCount, 1] as Range).Value2 && isCopyStartIndex)
                        {
                            copyStartIndex = rowCount;
                            Debug.WriteLine(departName + " 복사 시작 좌표값 : {0}", rowCount);
                            isCopyStartIndex = false;
                        }

                        if (departName == (string)(WholeInputDataRange.Cells[rowCount, 1] as Range).Value2 && isCopyStartIndex == false)
                        {
                            copyEndIndex = rowCount;
                        }
                    }

                    if (copyStartIndex == 0 || copyEndIndex == 0)
                    {
                        continue;
                    }
                    else
                    {
                        Debug.WriteLine(departName + " 복사 종료 좌표값 : {0}", copyEndIndex);
                        Debug.WriteLine(departName + " 복사 범위 값 : {0}", copyEndIndex - copyStartIndex + 1);

                        int nextEndIndex = nextStartIndex + copyEndIndex - copyStartIndex;

                        Debug.WriteLine(departName + " 기입 마지막 위치 : {0}", nextEndIndex);

                        var fromIndex = "A" + copyStartIndex.ToString() + ":Z" + copyEndIndex.ToString();
                        var toIndex = "A" + nextStartIndex.ToString() + ":Z" + nextEndIndex.ToString();

                        var from = InputDataWorksheet.Range[fromIndex];
                        var toWorksheet = OutputAllWorkbook.Worksheets.Item[collegename] as Worksheet;
                        var to = toWorksheet.Range[toIndex];

                        from.Copy(to);
                    }                    
                }
                Debug.WriteLine(targetWorksheet.UsedRange.Rows.Count);

                if (targetWorksheet.UsedRange.Rows.Count >= 2)
                {
                    StudentNum.Add(collegename, targetWorksheet.UsedRange.Rows.Count);
                }
            }            
            Debug.WriteLine("=============================단과대별 학과 분류하여 워크시트 데이터 기입 종료=============================");

            // ResultEachCollege();
        }

        Dictionary<string, int> ResultIndexDictionary = new Dictionary<string, int>();

        public void ResultEachCollege()
        {
            var collegelist = StudentNum.Keys;

            foreach(var college in collegelist)
            {
                Debug.WriteLine(college + " 결과 작성 시작");

                if(college == "전체인원")
                {
                    continue;
                }

                var targetworksheet = OutputAllWorkbook.Worksheets.Item[college] as Worksheet;

                // 총 단과대 인원 작성
                var studentcountindex = targetworksheet.UsedRange.Rows.Count + 3;

                ResultIndexDictionary.Add(college, studentcountindex);

                var writeplace = targetworksheet.Range[targetworksheet.Cells[studentcountindex, 5], targetworksheet.Cells[studentcountindex, 26]];
                
                writeplace.Value2 = StudentNum[college];

                // 개별 단과대 개별 이상자 인원수 파악
                studentcountindex -= 1; // 카운트 낱개 갯수 위치 조정

                // 마지막 끝나느 studentcountindex == 컬럼 갯수 위치
                for (var columnindex = 5;  columnindex < 27; columnindex++)
                {
                    var columncount = ColumnCounter(targetworksheet, columnindex, college);
                    (targetworksheet.Cells[studentcountindex, columnindex] as Range).Value2 = columncount;

                    studentcountindex += 2;

                    var input = Math.Round((float)columncount / StudentNum[college], 4);
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

                var inputPercentageValue = Math.Round((float)inputValue / StudentNum["전체인원"], 4);

                (graphworksheet.Cells[3, columnIndex] as Range).Value2 = inputPercentageValue;
            }
        }

        private int ColumnCounter(Worksheet targetworksheet, int columnindex, string college)
        {
            var count = 0;

            for(var index = 1; index < StudentNum[college]; index++)
            {
                var temp = Convert.ToInt32((targetworksheet.Cells[index, columnindex] as Range).Value2);

                if(temp >= 70)
                {
                    count++;
                }
            }

            return count;
        }

        public void MisfitFiltering()
        {
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
                    if ( paranoiaValue >= 60 || psychosisValue >= 60)
                    {
                        Debug.WriteLine("적성인식-예방 열정보 삽입");
                        preventAptitudeRecList.Add(rowCount);
                    }

                    // 스트레스
                    if ( depressedValue >= 60 || unrestValue >= 60)
                    {
                        Debug.WriteLine("스트레스-예방 열정보 삽입");
                        preventStressList.Add(rowCount);
                    }

                    // 외상경험
                    if ( ptsdValue >= 60 || addictionValue >= 60)
                    {
                        Debug.WriteLine("외상경험-예방 열정보 삽입");
                        preventTraumaList.Add(rowCount);
                    }

                    // 고립
                    if ( fearValue >= 60 || paranoiaValue >= 60)
                    {
                        Debug.WriteLine("고립-예방 열정보 삽입");
                        preventIsolateList.Add(rowCount);
                    }

                    // 대인갈등
                    if ( maniaValue >= 60 || angerValue >= 60)
                    {
                        Debug.WriteLine("대인갈등-예방 열정보 삽입");
                        preventIPConflictList.Add(rowCount);
                    }
                }


                // 문제집단
                if (targetValue >= 70)
                {
                    // 적성인식
                    if ( paranoiaValue >= 70 || psychosisValue >= 70)
                    {
                        Debug.WriteLine("적성인식-문제 열정보 삽입");
                        seriousAptitudeRecList.Add(rowCount);
                    }

                    // 스트레스
                    if ( depressedValue >= 70 || unrestValue >= 70)
                    {
                        Debug.WriteLine("스트레스-문제 열정보 삽입");
                        seriousStressList.Add(rowCount);
                    }

                    // 외상경험
                    if ( ptsdValue >= 70 || addictionValue >= 70)
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
                    if ( maniaValue >= 70 || angerValue >= 70)
                    {
                        Debug.WriteLine("대인갈등-문제 열정보 삽입");
                        seriousIPConflictList.Add(rowCount);
                    }
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

            currentRowIndex+=2;

            (targetWorksheet.Cells[currentRowIndex, departIndex] as Range).Value = "학과";
            (targetWorksheet.Cells[currentRowIndex, numIndex] as Range).Value = "학번";
            (targetWorksheet.Cells[currentRowIndex, nameIndex] as Range).Value = "성명";
            (targetWorksheet.Cells[currentRowIndex, sexIndex] as Range).Value = "성별";
            (targetWorksheet.Cells[currentRowIndex, target1Index] as Range).Value = (WholeInputDataRange.Cells[1, target1] as Range).Value2;
            (targetWorksheet.Cells[currentRowIndex, target2Index] as Range).Value = (WholeInputDataRange.Cells[1, target2] as Range).Value2;
            (targetWorksheet.Cells[currentRowIndex, stressIndex] as Range).Value = "스트레스취약성";

            foreach (var index in targetlist)
            {
                (targetWorksheet.Cells[currentRowIndex, departIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 1] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, numIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 2] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, nameIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 3] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, sexIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, 4] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, target1Index] as Range).Value2 = (WholeInputDataRange.Cells[index, target1Index] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, target2Index] as Range).Value2 = (WholeInputDataRange.Cells[index, target2Index] as Range).Value2;
                (targetWorksheet.Cells[currentRowIndex, stressIndex] as Range).Value2 = (WholeInputDataRange.Cells[index, stressIndex] as Range).Value2;

                currentRowIndex++;
            }

            return currentRowIndex;
        }
    }
}