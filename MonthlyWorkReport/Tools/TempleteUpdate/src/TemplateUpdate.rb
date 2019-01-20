# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/Define.rb"

# ==========================="
# src
# ==========================="
class TemplateUpdate
	private
	IN_ROOT						= File.dirname(__FILE__) + "/../in"
	OUT_ROOT 					= File.dirname(__FILE__) + "/../out"
	TEMPLATE_FILE_NAME	= File.dirname(__FILE__) + "/../../Template/1-UP作業月報_template.#{EXT_NAME}"

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
		date_range = Excel.calcRangeStr( STAFF_SHEET_DATA.DateStartColumn, STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, date_range, dst_ws, date_range )

		# 各種「曜日」
		say_week_range = Excel.calcRangeStr( STAFF_SHEET_DATA.DayWeekColumn, STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, say_week_range, dst_ws, say_week_range )

		# 各種「総勤務時間」
		work_time_range = Excel.calcRangeStr( STAFF_SHEET_DATA.WorkTimeColumn, STAFF_SHEET_START_ROW, work_rows )
		Excel.rangeCopyFast( src_ws, work_time_range, dst_ws, work_time_range )

		# 各種「合計」行
		total_time_range = "A#{work_rows}:L#{work_rows}"
		Excel.rangeCopyFast( src_ws, total_time_range, dst_ws, total_time_range )
		
		# セルをロック（編集不可）にしてシートを保護
		dst_ws.range( "#{FORMAT_STAFF_SHEET_CALENDAR}" ).Locked = true
		dst_ws.Protect
	end

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

	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )

		@file_list = Array.new()
		# [in] にある excel のファイルリストを作成
		Dir.glob( "#{IN_ROOT}" + "/**/" + "*.#{EXT_NAME}" ) do |file_path|
			@file_list.push( file_path )
		end

		if( @file_list.size() == 0 )
			assertLogPrintFalse( "in フォルダにファイルがありません" )
		end
	end

	def update()

		Excel.runDuring(false, false) do |excel|

			# コピーしたブックを開く
			fso = WIN32OLE.new('Scripting.FileSystemObject')
			wb_templete = excel.workbooks.open({'filename'=> fso.GetAbsolutePathName( TEMPLATE_FILE_NAME ), 'updatelinks'=> 0})
			ws_templete = wb_templete.worksheets( SHEET_NAME_TEMPLATE_DATA )

			# ファイルの数だけ
			@file_list.each { |file_path|
				wb_staff = excel.workbooks.open({'filename'=> fso.GetAbsolutePathName( file_path ), 'updatelinks'=> 0})
				ws_staff = wb_staff.worksheets( "#{getStaffName(file_path)}" )

				# フォーマットを更新
				excelFormatUpdate( ws_templete, ws_staff )

				# 更新したものをoutフォルダにセーブして閉じる
				out_path = file_path.gsub( "in", "out" )
				wb_staff.saveAs( fso.GetAbsolutePathName("#{out_path}") )
				wb_staff.close()

				# ログ用
				puts "update excel => #{File.basename( out_path )}"
			}
			wb_templete.close()
		end
	end
end
