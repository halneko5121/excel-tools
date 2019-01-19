# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"

# ==========================="
# src
# ==========================="
class TemplateExcelCreate
	private
	OUT_ROOT 				= File.dirname(__FILE__) + "/../../../Users"
	FILE_PREFIX				= "【1-UP】健康チェックリスト"
	TEMPLATE_FILE_NAME		= File.dirname(__FILE__) + "/../../Template/#{FILE_PREFIX}_templete.xlsx"
	DEAD_LINE_STR_PREFIX	= "XXXX"

	private
	#----------------------------------------------
	# @biref	出力先のパスを取得
	# @parm		id			社員番号
	# @parm		abbrev_name	社員略称
	# @parm		calendar	月報日時
	#----------------------------------------------
	def getOutputPath( id, abbrev_name, calendar )

		# 数値を3桁に変換
		staff_id	= "%03d" % id.to_i
        abbrev_name = abbrev_name.encode( Encoding::UTF_8 )
		file_name	= "#{FILE_PREFIX}_#{calendar.to_i.to_s}_#{staff_id}_#{abbrev_name}.xlsx".encode(Encoding::Windows_31J)
		out_path	= "#{OUT_ROOT}/#{file_name}"

		return out_path;
	end

	#----------------------------------------------
	# @biref	生成年月をDate変換
	# @parm		create_calendar	生成年月
	#----------------------------------------------
	def convertCreateDate( create_calendar )

		# 2013xx => [2013][xx]に分割
		str_calendar = splitYearMonth( create_calendar )

		# 年 / 月 を算出
		year	 			= str_calendar[0].to_i
		mouth				= str_calendar[1].to_i
        create_time			= Date.new( year, mouth )

		return create_time
	end

	#----------------------------------------------
	# @biref	生年月日から年齢を算出
	# @parm		birth_day		生年月日
	# @parm		create_calendar	生成年月
	#----------------------------------------------
	def calcStaffAge( birth_day, create_calendar )

		# 年 / 月 を算出
        time_now			= convertCreateDate( create_calendar )
        date_time_birth		= DateTime.parse( birth_day )
        time_birth			= Date.new( date_time_birth.year, date_time_birth.mon, date_time_birth.day )

		diff			= time_now - time_birth
		result_year		= diff.to_f / 365
		result_mouth	= (result_year - result_year.floor) * 12
		result_age		= "#{result_year.floor}歳 #{result_mouth.ceil}ヶ月"

		return result_age;
	end

	#----------------------------------------------
	# @biref	勤続年数を算出
	# @parm		joining_time	入社日
	# @parm		create_calendar	生成年月
	#----------------------------------------------
	def calcWorkTime( joining_time, create_calendar )

		# 年 / 月 を算出
        time_now			= convertCreateDate( create_calendar )
        date_time_joining	= DateTime.parse( joining_time )
        time_joining		= Date.new( date_time_joining.year, date_time_joining.mon, date_time_joining.day )

		work_time		= time_now - time_joining
		result_year		= work_time.to_f / 365
		result_mouth	= (result_year - result_year.floor) * 12
		result_work_time = "#{result_year.floor}年 #{result_mouth.ceil}ヶ月"

		return result_work_time;
	end

	#----------------------------------------------
	# @biref	「報告締切日」を算出
	# @parm		defult_str		デフォルトの文言
	# @parm		dead_line_str	置換する文言
	#----------------------------------------------
	def calcDeadLineString( defult_str, dead_line_str )

		# デフォルトの文字列の 「XXXX」 をパラメータで置換する
		return defult_str.gsub( "XXXX", dead_line_str )
	end

	#----------------------------------------------
	# @biref	パラメータの設定(各社員シート)
	# @parm		wb			設定を行うワークブック
	# @parm		param_hash	パラメータを格納したハッシュ
	#----------------------------------------------
	def setWsParamStaffSheet( wb, param_hash )

		# 社員名（フルネーム）
		ws_staff	  = wb.worksheets("テンプレート")
		ws_staff.name = "#{param_hash[:abbrev_name]}"

		ws_staff.UnProtect
			# 数値を3桁に変換
			ws_staff.Cells.Item(6, 3).Value = "%03d" % "#{param_hash[:id]}".to_i
			ws_staff.Cells.Item(7, 3).Value = "#{param_hash[:name]}"
			ws_staff.Cells.Item(8, 3).Value = "#{param_hash[:job_type]}"

			ws_staff.Cells.Item(6, 7).Value = "#{param_hash[:gender]}"
			ws_staff.Cells.Item(7, 7).Value = "#{param_hash[:age]}"#calcStaffAge( "#{param_hash[:birth_day]}", "#{param_hash[:create_calendar]}" )
			ws_staff.Cells.Item(8, 7).Value = "#{param_hash[:joining_time]}"#calcWorkTime( "#{param_hash[:joining_time]}", "#{param_hash[:create_calendar]}" )

			ws_staff.Cells.Item(7, 10).Value = "#{param_hash[:last_month_over_time]}"
			ws_staff.Cells.Item(8, 10).Value = "#{param_hash[:last_month_over_time2]}"
			ws_staff.Cells.Item(73, 1).Value = calcDeadLineString( ws_staff.Cells.Item(73, 1).Value, "#{param_hash[:report_dead_line]}" )

			dead_line_str = ws_staff.Cells.Item(73, 1).Value
			Excel.setStringColor( ws_staff, 73, 1, dead_line_str, "#{param_hash[:report_dead_line]}" )

		ws_staff.Protect

		# 最初のセルをアクティブ
		ws_staff.Range("A1").Activate

		#　シート保護をしない or マクロ有りブックにする
#		ws_staff.EnableOutlining = true
#		ws_staff.Protect( {'Contents' => true} )
#		ws_staff.Protect( {'UserInterfaceOnly' => true} )
#		ws_staff.Protect( {'allowformattingcells' => true} )
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

		# 出力OKのものだけ出力する
		result_staff_list = Array.new
		staff_list.each{|data|
			utf_8_is_output = data[:is_output].encode( Encoding::UTF_8 )
			if ( utf_8_is_output.index( "◯" ) != nil)
				result_staff_list.push(data)
			end
		}

		puts "excel count = #{result_staff_list.size()}"

		Excel.runDuring(false, false) do |excel|

			fso = WIN32OLE.new('Scripting.FileSystemObject')

			# 社員数だけ
			staff_number = 1
			result_staff_list.each{|data|

				# 出力先のパスを取得
				out_path = getOutputPath( "#{data[:id]}", "#{data[:abbrev_name]}", "#{data[:create_calendar]}" )

				# テンプレートのブックをコピー
				fso.CopyFile( "#{TEMPLATE_FILE_NAME}", fso.GetAbsolutePathName( out_path ) )

				# コピーしたブックを開く
				wb = excel.workbooks.open({'filename'=> fso.GetAbsolutePathName( out_path ), 'updatelinks'=> 0})
				excel.displayAlerts = false
				excel.visible		= false

				# パラメータの設定
				# 左上をアクティブにしてスクロールしておく
				setWsParamStaffSheet( wb, data )
				excel.ActiveWindow.ScrollRow = 1
				excel.ActiveWindow.ScrollColumn = 1

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
