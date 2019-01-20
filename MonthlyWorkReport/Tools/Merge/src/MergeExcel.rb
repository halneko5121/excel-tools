# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/Define.rb"

# ==========================="
# class
# ==========================="
class MergeExcel
	private
	# Const
	OUT_ROOT 				= "."
	TEMPLATE_FILE_NAME		= File.expand_path(File.dirname(__FILE__)) + "/../../Template/Template.#{EXT_NAME}"
	CHECK_DIR				= File.expand_path(File.dirname(__FILE__)) + "/../../../Users"
	SEARCH_FILE 			= "*1-UP*.#{EXT_NAME}"
	START_ROW_PRORATED		= 5
	CHECK_DATA_RANGE		= 31
	START_ROW_CHECK_DATA	= 3

	def setFileList( list )

		# パスの変換(\\ => /)
		check_dir = CHECK_DIR.gsub( "\\", "/" )
		dbgPuts "check_dir = #{check_dir}"

		# パターンにマッチするファイルパスを追加
		Dir.glob( "#{check_dir}" + "/**/" + "#{SEARCH_FILE}" ) do |file_path|
			list.push( file_path )
		end

		# ascii順に並び替え
		list.sort!

		puts "excel count = #{list.size()}"

		if( list.size == 0 )
			puts "=============== error ==============="
			puts "not found merging files!!"
			puts "Users フォルダにファイルがあるかお確かめ下さい"
			puts "=============== error ==============="
			exit()
		end
	end

	#----------------------------------------------
	# @biref	ファイルパスから、出力先のパスを取得
	# @parm		file_name	作業月報名
	# @parm		calender	月報日時
	#----------------------------------------------
	def getOutputPath( file_path )

		file_name		= File.basename( file_path )
		file_name_info	= file_name.split( "_" )
		out_file_name	= file_name_info[2]
		calender		= file_name_info[3]
		output_path 	= "#{OUT_ROOT}/#{out_file_name}_#{calender}"
		return ( output_path.gsub( ".xlsm", ".xlsx" ) )
	end

	#----------------------------------------------
	# @biref	システムシートへ値のコピー
	# @parm		system_ws	システムシート（開始データ、未定）
	# @parm		src_ws		コピー元のシート
	#----------------------------------------------
	def setWsParamSystemSheet( system_ws, src_ws )
		system_ws.UnProtect

		#　年/月/期間
		system_ws.range( "#{FORMAT_STAFF_SHEET_CALENDAR}" ).Value	= src_ws.range( "#{FORMAT_STAFF_SHEET_CALENDAR}" ).Value
		#　日付
		system_ws.range( "#{FORMAT_STAFF_SHEET_DATE}" ).Value		= src_ws.range( "#{FORMAT_STAFF_SHEET_DATE}" ).Value
		#　曜日
		system_ws.range( "#{FORMAT_STAFF_SHEET_DAY_WEEK}" ).Value	= src_ws.range( "#{FORMAT_STAFF_SHEET_DAY_WEEK}" ).Value
		#　各種列名設定
		system_ws.range( "#{FORMAT_STAFF_SHEET_PROJECT}" ).Value	= src_ws.range( "#{FORMAT_STAFF_SHEET_PROJECT}" ).Value

		system_ws.Protect
	end

	#----------------------------------------------
	# @biref	システムシートへ値のコピー
	# @parm		system_ws	システムシート（開始データ、未定）
	#----------------------------------------------
	def setWsParamWorksSheet( system_ws )

		system_ws.UnProtect
		system_ws.range( "#{CELL_START_PROJECT_SUM}:#{CELL_START_PROJECT_SUM}" ).Value = "=SUM(開始データ:未定!#{CELL_START_PROJECT_SUM}:#{CELL_START_PROJECT_SUM})"
		system_ws.range( "#{CELL_START_PROJECT_SUM}:#{CELL_START_PROJECT_SUM}" ).copy
		system_ws.range( "#{FORMAT_WORKS_SHEET_PROJECT_SUM}" ).pastespecial

		range_check_project		= "D#{ROW_START_PROJECT_SUM + 1}"
		range_check_project_sum	= "D#{ROW_START_PROJECT_SUM + 2}:D#{ROW_START_PROJECT_SUM + 2}"
		system_ws.range( "#{range_check_project_sum}" ).Value = "=IF(#{range_check_project}=SUM(開始データ:未定!#{range_check_project}:#{range_check_project}),\"OK\",\"ちがう\")"
		system_ws.range( "#{range_check_project_sum}" ).copy
		system_ws.range( "#{FORMAT_WORKS_SHEET_CHECK_SUM}" ).pastespecial

		# セルをロック（編集不可）にしてシートを保護
		system_ws.range( "#{FORMAT_WORKS_SHEET_PROJECT_SUM}" ).Locked = true
		system_ws.range( "#{FORMAT_WORKS_SHEET_CHECK_SUM}" ).Locked = true
		system_ws.Protect
	end

	#----------------------------------------------
	# @biref	ファイルパスからユーザー名を探します
	# @parm		wb			設定を行うワークブック
	#----------------------------------------------
	def searchUserName( file_path )

		# [1] => 社員名略称（シート名）
		file_name		= File.basename( file_path )
		file_name_info	= file_name.split( "_" )
		user_name		= file_name_info[1]
		return user_name
	end

	#----------------------------------------------
	# @biref	区分別按分表シートの設定
	# @parm		wb			設定を行うワークブック
	#----------------------------------------------
	def setWsParamProratedTable( wb )

		# 略称（ファイル名/シート名）
		ws_propateed = wb.worksheets("#{SHEET_NAME_PRORATED_TABLE}")
		ws_propateed.Cells.Columns(22).Hidden = true
		ws_propateed.Cells.Columns(23).Hidden = true
	end

	#----------------------------------------------
	# @biref	届書チェックシートのマージ
	# @parm		excel		Excel クラス
	# @parm		wb_merge	マージするワークブック
	#----------------------------------------------
	def mergeCheckSheet( wb_src, wb_dst, count )

		range_start		= START_ROW_CHECK_DATA		+ CHECK_DATA_RANGE * count
		range_end		= (START_ROW_CHECK_DATA-1)	+ CHECK_DATA_RANGE * (count + 1)
		ws_src_check	= wb_src.worksheets("#{SHEET_NAME_CHECK}")
		ws_dst_check	= wb_dst.worksheets("#{SHEET_NAME_CHECK}")
		ws_src_check.range( "#{RANGE_CHECK_VALUE}" ).copy
		ws_dst_check.range( "B#{range_start}:X#{range_end}" ).pastespecial
	end

	#----------------------------------------------
	# @biref	各スタッフエクセルからシートの設定
	# @parm		excel		Excel クラス
	# @parm		wb_merge	マージするワークブック
	#----------------------------------------------
	def setWsParamEachStaffSheet( excel, wb_merge )

		# 「区分別按分表」シートを取得
		ws_dst_department = wb_merge.worksheets("#{SHEET_NAME_PRORATED_TABLE}")

		ws_start			= wb_merge.worksheets("#{SHEET_NAME_TEMPLATE_DATA}")
		start_sheet_number	= ws_start.index

		# HITしたファイルの数だけ
		count = 0
		@file_path_list.each{|file_path|

			puts "merge => #{File.basename( "#{file_path}" )}"

			# excel Open
			wb_staff = excel.workbooks.open({'filename'=> "#{file_path}"})

			# 「各々の作業月報」シートをコピー
			user_name = searchUserName( file_path )
			Excel.sheetCopy( wb_staff, "#{user_name}", wb_merge, start_sheet_number+count)

			# 届け出チェックシートの列を削除
			if( IS_CHECK_SHEET_MIX == true )
				ws_staff = wb_merge.worksheets("#{user_name}")
				ws_staff.unprotect
				ws_staff.range("A:Z").delete
				ws_staff.protect
			end

			# 「区分別按分表」シートに行を挿入してコピー・ペースト
			ws_src_department	= wb_staff.worksheets("#{SHEET_NAME_PRORATED_TABLE}")
			paste_row			= (START_ROW_PRORATED+1) + count # 開始データの分を残しておく
			Excel.rowCopyAndInsert( ws_src_department, START_ROW_PRORATED, ws_dst_department, paste_row )
			ws_dst_department.Cells.Item(paste_row, 23).Value = 0# チェックシーシート分の列指定を0に

			if( IS_CHECK_SHEET_MIX == false )
				# 「届書チェック」シートの値をマージ
				mergeCheckSheet( wb_staff, wb_merge, count )
			end

			wb_staff.Application.CutCopyMode = false
			wb_staff.close(0)
			count += 1
		}

		# 「開始データ」さんの行は削除しておく
		ws_dst_department.range("#{START_ROW_PRORATED}:#{START_ROW_PRORATED}").delete
		excel.screenUpdating = true
		ws_dst_department.Activate
		excel.range("C2").Select
		excel.screenUpdating = false
	end

	#----------------------------------------------
	# @biref	マージexcel にパラメータの適応を行う
	# @parm		wb_merge	マージするワークブック
	# @parm		excel		Excel クラス
	#----------------------------------------------
	def applyParamMergeWb( wb_merge, data )

		#パラメータの取得
		is_protected		= "#{data[:is_protected]}"			#"シートを保護するか"
		is_delete_ws_check	= "#{data[:is_delete_ws_check]}"	#"[届け出チェックシート]を削除するか"

		# パラメータの適応:シートの保護
		if( is_protected == "true" )
			wb_merge.worksheets("#{SHEET_NAME_PRORATED_TABLE}").Protect
		else
			ws_count = wb_merge.worksheets.Count
			ws_count = ws_count.to_i
			for index in 1..ws_count do
				wb_merge.worksheets( index ).Unprotect
			end
		end

		# パラメータの適応:チェックシートの削除
		if( is_delete_ws_check == "true" )
			wb_merge.worksheets("#{SHEET_NAME_CHECK}").delete
#			wb_merge.worksheets("#{SHEET_NAME_CHECK_DESCRIPT}").delete
		end
	end

	#----------------------------------------------
	# @biref	マージexcel に必要なシートを設定する
	# @parm		excel		Excel クラス
	# @parm		merge_wb	マージするワークブック
	# @parm		tamplete_wb	テンプレートワークブック
	#----------------------------------------------
	def setDefaultSheetMergeExcel( excel, merge_wb, tamplete_wb )

		merge_wb_ws_count = merge_wb.worksheets.Count

		sheet_number = 1

		# 「注意」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_WARNING}", merge_wb, sheet_number)
		sheet_number += 1

		# 「テンプレート」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_TEMPLATE_DATA}", merge_wb, sheet_number)
		sheet_number += 1

		# 「未定」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_ANOTHER}", merge_wb, sheet_number)
		sheet_number += 1

		# 「届書チェックシート」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_CHECK}",merge_wb, sheet_number)
		sheet_number += 1

		# 「業務別月報」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_WORKS}", merge_wb, sheet_number)
		sheet_number += 1

		# 「区分別按分表」シートをコピー
		Excel.sheetCopy( tamplete_wb, "#{SHEET_NAME_PRORATED_TABLE}", merge_wb, sheet_number)
		sheet_number += 1

		# 最初に生成されたシートを削除
		excel.displayAlerts = false
		for index in 1..merge_wb_ws_count do
			merge_wb.worksheets("Sheet#{index}").delete()
		end
	end

	public
	def initialize()
		@file_path_list = Array.new
	end

	def main( param_hash )

		# ファイルパスをリストに設定
		setFileList( @file_path_list )

		# excelの処理
		Excel.runDuring(false, false) do |excel|

			# マージ用 excel を開く
			wb = excel.workbooks.open({'filename'=> "#{TEMPLATE_FILE_NAME}", 'updatelinks'=> 0})

			# 新規ワークブックを作成・必要なシートをコピー
			wb_merge = excel.workbooks.add()
			setDefaultSheetMergeExcel( excel, wb_merge, wb )
			wb.close(0)

			# スタッフの数だけループ
			param_hash.each{|data|

				# 各々のスタッフ　excel　の値を設定
				setWsParamEachStaffSheet( excel, wb_merge )

				# 届け出チェックシートの列を削除
				if( IS_CHECK_SHEET_MIX == true )
					ws_start = wb_merge.worksheets("#{SHEET_NAME_TEMPLATE_DATA}")
					ws_start.unprotect
					ws_start.range("A:Z").delete
					ws_start.protect
				end

				# 「開始データ」「未定」「業務別月報」に値をコピー(年月/期間)
				ws_staff	= wb_merge.worksheets("門井")
				ws_start	= wb_merge.worksheets("#{SHEET_NAME_TEMPLATE_DATA}")
				ws_no_fixed = wb_merge.worksheets("#{SHEET_NAME_ANOTHER}")
				ws_monthly	= wb_merge.worksheets("#{SHEET_NAME_WORKS}")

				setWsParamSystemSheet(ws_no_fixed, ws_staff)
				setWsParamSystemSheet(ws_monthly, ws_staff)
				setWsParamSystemSheet(ws_start, ws_staff)

				# 「業務別月報」の設定
				setWsParamWorksSheet(ws_monthly)

				# [開始データ] シートは非表示にしておく
				ws_start.visible = false

				# 「区分別按分表シート」の設定
				setWsParamProratedTable( wb_merge )

				# パラメータの適応
				applyParamMergeWb( wb_merge, data )

				# ファイル名から一部を拝借して設定
				out_path = getOutputPath( @file_path_list[0] )

				# 最初のシートをアクティブにして終了
				wb_merge.worksheets(1).Activate

				# 保存して閉じる
				fso	= WIN32OLE.new('Scripting.FileSystemObject')
				wb_merge.saveAs( fso.GetAbsolutePathName("#{out_path}") )
				wb_merge.close(0)

				puts "output merge excel = #{out_path}"
			}
		end
	end
end
