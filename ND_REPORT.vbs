' 各種設定
const RESULT_DIR = "C:\github\ND_trial\result\"
const DB_FILE_PATH = "C:\github\ND_trial\data\DB.xlsx"
const DB_SHEET_NAME = "CRF回収"
const FACILITY_COL = 5
const START_ROW = 3
const END_ROW = 30

' エクセルオブジェクト
set XLS = CreateObject("Excel.Application")
XLS.Visible = true

'DBのエクセルブックをオープンする
set dbBook = XLS.workbooks.open(DB_FILE_PATH)
set dbSheet = dbBook.sheets(DB_SHEET_NAME)

'施設リストを作成する
Set facilities = CreateObject("Scripting.Dictionary")
for y = START_ROW to END_ROW
	facility = dbSheet.cells(y, FACILITY_COL)
	if not facilities.exists(facility) then
		facilities.Add facility, facility
	end if
next

'施設リストから施設を順番に読み取る
for each facility in facilities

	'ブックを作成
	set NEWBOOK	= XLS.Workbooks.Add
	NEWBOOK.Application.DisplayAlerts = False

	'不要なシートを削除
	orgSheetCnt = NEWBOOK.Sheets.count
	For i = 1 to orgSheetCnt-1
		NEWBOOK.Sheets(1).delete
	Next

	'施設名で保存する
	NEWBOOK.saveAs(RESULT_DIR & facility & ".xlsx")
next

dbBook.close
XLS.quit
