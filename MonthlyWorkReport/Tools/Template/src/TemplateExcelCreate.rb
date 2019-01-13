# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require "fileutils"
require "find"
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/Define.rb"

# ==========================="
# src
# ==========================="
class TemplateExcelCreate
	private
	OUT_ROOT 					= File.dirname(__FILE__) + "/../../../Users"
	TEMPLATE_FILE_NAME	= File.dirname(__FILE__) + "/../../Template/1-UP作業月報_template.#{EXT_NAME}"

	private
	#----------------------------------------------
	# @biref	出力先のパスを取得
	# @parm		number		データ番号
	# @parm		abbrev_name	社員略称
	# @parm		calendar	月報日時
	#----------------------------------------------
	def getOutputPath( number, abbrev_name, calendar )
	
		# 数値を3桁に変換
		staff_num	= "%03d" % number

        abbrev_name= abbrev_name.encode( Encoding::UTF_8 ) 
		file_name	= "#{staff_num}_#{abbrev_name}_1-UP作業月報_#{calendar}.#{EXT_NAME}".encode(Encoding::Windows_31J)
		out_path	= "#{OUT_ROOT}/#{file_name}"
		
		return out_path;
	end

	#----------------------------------------------
	# @biref	西暦を平成に変換する
	#----------------------------------------------
	def getYearNumber( year )

		# year を 1文字 / 3文字に分割 => 下三桁に12を加算
		year_number_array = year.unpack("a1a3")
		return ( year_number_array[1].to_i + 12 )
	end

	#----------------------------------------------
	# @biref	パラメータの設定(各社員シート)
	# @parm		wb			設定を行うワークブック
	# @parm		param_hash	パラメータを格納したハッシュ
	#----------------------------------------------
	def setWsParamStaffSheet( wb, param_hash )

		# 社員名（フルネーム）
		ws_staff	  = wb.worksheets("開始データ")
		ws_staff.name = "#{param_hash[:abbrev_name]}"
		ws_staff.Cells.Item(2, 10+ADD_CULMNS_CHECK_SHEET).Value = "#{param_hash[:name]}"	# 氏名
		
		# 2013xx => [2013][xx]に分割
		str_calendar = getSplitCalendar("#{param_hash[:create_calendar]}")

		# 平成年 / 月 を算出
#		year_number = getYearNumber( str_calendar[0] )
#		year_number = "#{str_calendar[0]}/#{str_calendar[1]}/1" => 平成歴の場合(セルの初期設定を"ggge"年にする）
		year_number = str_calendar[0].to_i
		mouth		= str_calendar[1].to_i

		ws_staff.Cells.Item(2, 5+ADD_CULMNS_CHECK_SHEET).Value = year_number
		ws_staff.Cells.Item(2, 6+ADD_CULMNS_CHECK_SHEET).Value = mouth
		ws_staff.Cells.Item(2, 7+ADD_CULMNS_CHECK_SHEET).Value = "#{param_hash[:period]}"	# 期間

		# セルをロック（編集不可）にしてシートを保護
		ws_staff.Cells.Item(2, 5+ADD_CULMNS_CHECK_SHEET).Locked = true
		ws_staff.Cells.Item(2, 6+ADD_CULMNS_CHECK_SHEET).Locked = true
		ws_staff.Cells.Item(2, 7+ADD_CULMNS_CHECK_SHEET).Locked = true
		if( IS_CHECK_SHEET_MIX == true )
			ws_staff.range("A:Z").Locked = false
		end
		ws_staff.Protect
		
		#　シート保護をしない or マクロ有りブックにする
#		ws_staff.EnableOutlining = true
#		ws_staff.Protect( {'Contents' => true} )
#		ws_staff.Protect( {'UserInterfaceOnly' => true} )
#		ws_staff.Protect( {'allowformattingcells' => true} )
	end

	#----------------------------------------------
	# @biref	パラメータの設定(届書シート)
	# @parm		wb			設定を行うワークブック
	# @parm		param_hash	パラメータを格納したハッシュ
	#----------------------------------------------
	def setWsParamCheckSheet( wb, param_hash )
	
		# 届書チェックシートの設定
		ws_check = wb.worksheets("#{SHEET_NAME_CHECK}")
		ws_check.Range("B3:B33").Value = "#{param_hash[:abbrev_name]}"
	end

	#----------------------------------------------
	# @biref	パラメータの設定(区分別按分表シート)
	# @parm		wb			設定を行うワークブック
	# @parm		param_hash	パラメータを格納したハッシュ
	#----------------------------------------------
	def setWsParamProratedTable( wb, param_hash )

		# 略称（ファイル名/シート名）
		ws_propateed = wb.worksheets("#{SHEET_NAME_PRORATED_TABLE}")
		ws_propateed.Cells.Item(5, 22).Value	= "#{param_hash[:abbrev_name]}"
		ws_propateed.Cells.Item(5, 23).Value	= ADD_CULMNS_CHECK_SHEET		
		ws_propateed.Cells.Item(6, 1).Value		= ws_propateed.Cells.Item(5, 1).Value
		ws_propateed.Cells.Item(6, 22).Value	= "#{param_hash[:abbrev_name]}"
		ws_propateed.Cells.Columns(22).Hidden	= true
		ws_propateed.Cells.Columns(23).Hidden	= true

		# 入社時期
		if( param_hash[:joining_time] != nil )
			ws_propateed.Cells.Item(5, 1).AddComment("#{param_hash[:joining_time]}")
		end
	end
	
	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )
	end

	def createExcel( staff_list )
	
		# ファイルが存在していた場合はファイルを削除
		Dir.glob( "#{OUT_ROOT}" + "/**/" + "*.*" ) do |file_path|
			File.delete "#{file_path}"
		end

		puts "excel count = #{staff_list.size()}"	

		Excel.runDuring(false, false) do |excel|

			fso = WIN32OLE.new('Scripting.FileSystemObject')

			# 社員数だけ
			staff_number = 1
			staff_list.each{|data|

				# 出力先のパスを取得
				out_path = getOutputPath( staff_number, "#{data[:abbrev_name]}", "#{data[:create_calendar]}" )

				# テンプレートのブックをコピー
				fso.CopyFile( "#{TEMPLATE_FILE_NAME}", fso.GetAbsolutePathName( out_path ) )

				# コピーしたブックを開く
				wb = excel.workbooks.open({'filename'=> fso.GetAbsolutePathName( out_path ), 'updatelinks'=> 0})
				excel.displayAlerts = false
				excel.visible		= false

				# シートを削除
				wb.worksheets( "業務別月報" ).delete()
				wb.worksheets( "未定" ).delete()
				if( IS_CHECK_SHEET_MIX == true )
					wb.worksheets( "届書チェックシート" ).delete()
				end
				
				# パラメータの設定
				setWsParamStaffSheet( wb, data )
				setWsParamProratedTable( wb, data )
				
				# シートは非表示にしておく  / [区分別按分表][届書チェック]
				wb.worksheets("#{SHEET_NAME_PRORATED_TABLE}").visible = false
				
				# セーブして閉じる
				wb.save()
				wb.close(0)
				
				staff_number += 1
				
				# ログ用
				puts "create excel => #{File::basename( out_path )}"
			}
		end
	end
end