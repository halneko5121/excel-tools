# -*- coding: utf-8 -*-

# ==========================="
# const
# ==========================="
SHEET_NAME_TEMPLATE_DATA	= "開始データ"

# excel のテンプレートフォーマット(スタッフシート)
FORMAT_STAFF_SHEET_CALENDAR	= "E2:G2"	#　「年/月/期間」のセル範囲
FORMAT_STAFF_SHEET_ALL		= "A1:O4"	# フォーマットいろいろの各種範囲
STAFF_SHEET_START_ROW 		= 5			# 入力データの開始行

StaffSheet = Struct.new(
				"FormatStaffSheet",
				:DateStartColumn,
				:DayWeekColumn,
				:WorkTimeColumn)
STAFF_SHEET_DATA = StaffSheet.new( "A", "B", "C" )

IS_CHECK_SHEET_MIX			= false	# 届書チェックシートを1シートにまとめる形。

if( IS_CHECK_SHEET_MIX == true )
	ADD_CULMNS_CHECK_SHEET	= 26
	EXT_NAME				= "xlsm"
else
	ADD_CULMNS_CHECK_SHEET	= 0
	EXT_NAME				= "xlsx"
end