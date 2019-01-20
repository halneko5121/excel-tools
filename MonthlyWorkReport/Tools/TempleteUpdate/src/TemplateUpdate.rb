# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + "/../../lib/excel.rb" )
require File.expand_path( File.dirname(__FILE__) + "/../../lib/util.rb" )

# ==========================="
# src
# ==========================="
class TemplateUpdate
	private
	IN_ROOT						= File.dirname(__FILE__) + "/../in"
	OUT_ROOT 					= File.dirname(__FILE__) + "/../out"
	TEMPLATE_FILE_NAME			= File.dirname(__FILE__) + "/../../Template/Template.xlsx"
	SHEET_NAME_TEMPLATE_DATA	= "開始データ"
	# excel のテンプレートフォーマット(スタッフシート)
	FORMAT_STAFF_SHEET_CALENDAR	= "E2:G2"	# 「年/月/期間」のセル範囲
	FORMAT_STAFF_SHEET_ALL		= "A1:O4"	# フォーマットいろいろの各種範囲
	STAFF_SHEET_START_ROW 		= 5			# 入力データの開始行

	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )

		# [in] にある excel のファイルリストを作成
		pattern_array =[ "*.xlsx" ]
		@file_list = getSearchFileList("#{IN_ROOT}", pattern_array)
		if( @file_list.size() == 0 )
			assertLogPrintFalse( "in フォルダにファイルがありません" )
		end
	end

	def update()

		Excel.runDuring(false, false) do |excel|

			# コピーしたブックを開く
			wb_templete = Excel.openWb( excel, TEMPLATE_FILE_NAME )
			ws_templete = wb_templete.worksheets( SHEET_NAME_TEMPLATE_DATA )

			# ファイルの数だけ
			@file_list.each { |file_path|
				wb_staff = Excel.openWb( excel, file_path )
				ws_staff = wb_staff.worksheets( "#{getStaffName(file_path)}" )

				# フォーマットを更新
				excelFormatUpdate( ws_templete, ws_staff )

				# 更新したものをoutフォルダにセーブして閉じる
				out_path = file_path.gsub( "in", "out" )
				Excel.saveAndClose( wb_staff , out_path )

				# ログ用
				puts "update excel => #{File.basename( out_path )}"
			}
			wb_templete.close()
		end
	end

	#----------------------------------------------
	# @biref	src excel のセルを dst にコピー
	#----------------------------------------------
	def excelFormatUpdate( src_ws, dst_ws )

		dst_ws.UnProtect

		# 「年/月/期間/氏名」のセル範囲を一時保存
		temp_year		= Excel.getCellValue( dst_ws, 2, 5 )
		temp_month		= Excel.getCellValue( dst_ws, 2, 6 )
		temp_period		= Excel.getCellValue( dst_ws, 2, 7 )
		temp_staff_name	= Excel.getCellValue( dst_ws, 2, 10 )

		# 各種フォーマット更新
		Excel.rangeCopyFast( src_ws, FORMAT_STAFF_SHEET_ALL, dst_ws, FORMAT_STAFF_SHEET_ALL )

		# 「年/月/期間/氏名」のセル範囲を設定
		setCalendarData( dst_ws, temp_year, temp_month,temp_period, temp_staff_name )

		# 行数を算出
		work_rows = Excel.getRow( src_ws, "計", 1)

		# 各種「日付」
		date_range = Excel.calcRangeStr( "A", STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, date_range, dst_ws, date_range )

		# 各種「曜日」
		say_week_range = Excel.calcRangeStr( "B", STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, say_week_range, dst_ws, say_week_range )

		# 各種「総勤務時間」
		work_time_range = Excel.calcRangeStr( "C", STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, work_time_range, dst_ws, work_time_range )

		# 各種「合計」行
		total_time_range = "A#{work_rows}:L#{work_rows}"
		Excel.rangeCopyFast( src_ws, total_time_range, dst_ws, total_time_range )

		# セルをロック（編集不可）にしてシートを保護
		dst_ws.range( "#{FORMAT_STAFF_SHEET_CALENDAR}" ).Locked = true
		dst_ws.Protect
	end

	private
	#----------------------------------------------
	# @biref	カレンダー情報　「年/月/期間」
	#----------------------------------------------
	def setCalendarData( dst_ws, year, month, period, staff_name )

		dst_ws.Cells.Item( 2, 5 ).Value = year
		dst_ws.Cells.Item( 2, 6 ).Value = month
		dst_ws.Cells.Item( 2, 7 ).Value = period
		dst_ws.Cells.Item( 2, 10 ).Value = staff_name
	end

	#----------------------------------------------
	# @biref	ファイルのパスから、社員名を取得
	#----------------------------------------------
	def getStaffName( file_path )

		file_name		= File.basename( file_path )
		file_name_info	= file_name.split( "_" )
		staff_name		= file_name_info[1]
		return staff_name
	end
end
