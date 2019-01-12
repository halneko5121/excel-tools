# -*- coding: utf-8 -*-

# ==========================="
# const
# ==========================="
SHEET_NAME_TEMPLATE_DATA	= "日報"
IS_CHECK_SHEET_MIX				= false	# 届書チェックシートを1シートにまとめる形

if( IS_CHECK_SHEET_MIX == true )
	ADD_CULMNS_CHECK_SHEET	= 26
	EXT_NAME				= "xlsm"
else
	ADD_CULMNS_CHECK_SHEET	= 0
	EXT_NAME				= "xlsx"
end
