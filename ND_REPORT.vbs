' 各種設定
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
	msgbox facility
next

dbBook.close
XLS.quit