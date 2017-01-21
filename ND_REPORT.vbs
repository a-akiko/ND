' �e��ݒ�
const RESULT_DIR = "C:\github\ND_trial\result\"
const DB_FILE_PATH = "C:\github\ND_trial\data\DB.xlsx"
const DB_SHEET_NAME = "CRF���"
const FACILITY_COL = 5
const START_ROW = 3
const END_ROW = 30

' �G�N�Z���I�u�W�F�N�g
set XLS = CreateObject("Excel.Application")
XLS.Visible = true

'DB�̃G�N�Z���u�b�N���I�[�v������
set dbBook = XLS.workbooks.open(DB_FILE_PATH)
set dbSheet = dbBook.sheets(DB_SHEET_NAME)

'�{�݃��X�g���쐬����
Set facilities = CreateObject("Scripting.Dictionary")
for y = START_ROW to END_ROW
	facility = dbSheet.cells(y, FACILITY_COL)
	if not facilities.exists(facility) then
		facilities.Add facility, facility
	end if
next

'�{�݃��X�g����{�݂����Ԃɓǂݎ��
for each facility in facilities

	'�u�b�N���쐬
	set NEWBOOK	= XLS.Workbooks.Add
	NEWBOOK.Application.DisplayAlerts = False

	'�s�v�ȃV�[�g���폜
	orgSheetCnt = NEWBOOK.Sheets.count
	For i = 1 to orgSheetCnt-1
		NEWBOOK.Sheets(1).delete
	Next

	'�{�ݖ��ŕۑ�����
	NEWBOOK.saveAs(RESULT_DIR & facility & ".xlsx")
next

dbBook.close
XLS.quit
