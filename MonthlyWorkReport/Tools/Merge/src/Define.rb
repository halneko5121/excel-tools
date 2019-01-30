# -*- coding: utf-8 -*-

# ==========================="
# const
# ==========================="
SHEET_NAME_WARNING			= "注意"
SHEET_NAME_TEMPLATE_DATA	= "開始データ"
SHEET_NAME_PRORATED_TABLE	= "区分別按分表"
SHEET_NAME_WORKS			= "業務別月報"
SHEET_NAME_ANOTHER			= "未定"

# excel のテンプレートフォーマット(スタッフシート)
FORMAT_STAFF_SHEET_CALENDAR	= "E2:G2"	# 「年/月/期間」のセル範囲
FORMAT_STAFF_SHEET_DAY_WEEK = "B5:B35"	# 「曜日」のセル範囲
FORMAT_STAFF_SHEET_DATE		= "A5:A35"	# 「日付」のセル範囲
FORMAT_STAFF_SHEET_PROJECT	= "D3:L4"	# 各種プロジェクト名のセル範囲
FORMAT_STAFF_SHEET_ALL		= "A1:O4"	# フォーマットいろいろの各種範囲

# excel のテンプレートフォーマット(作業別月報)
ROW_START_PROJECT_SUM			= 35		# 各種プロジェクトごとの作業時間の開始行
CELL_START_PROJECT_SUM			= "D5"		# 各種プロジェクトごとの作業時間の開始セル
FORMAT_WORKS_SHEET_PROJECT_SUM	= "D5:L35"	# 各種プロジェクトごとの作業時間のセル範囲
FORMAT_WORKS_SHEET_CHECK_SUM	= "C37:L37"	# 各種プロジェクトごとの作業時間判定のセル範囲

# excel のテンプレートフォーマット(届書チェックシート)
RANGE_CHECK_VALUE				= "B3:X33"
