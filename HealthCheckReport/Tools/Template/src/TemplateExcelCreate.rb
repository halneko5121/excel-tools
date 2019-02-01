# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + "/../../lib/Excel.rb" )
require File.expand_path( File.dirname(__FILE__) + "/../../lib/Util.rb" )

# ==========================="
# src
# ==========================="
class TemplateExcelCreate
	private
	OUT_ROOT 					= File.expand_path( File.dirname(__FILE__) + "/../../../Users" )
	FILE_PREFIX					= "【1-UP】健康チェックリスト"
	TEMPLATE_FILE_NAME			= File.expand_path( File.dirname(__FILE__) + "/../../Template/Template.xlsx" )
	SHEET_NAME_TEMPLATE_DATA	= "テンプレート"
	DEAD_LINE_STR_PREFIX		= "XXXX"

	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )
	end

	def execute( staff_list )

		# ファイルが存在していた場合はファイルを削除
		pattern = [ "*.xlsx" ]
		allRemoveFile("#{OUT_ROOT}", pattern)

		# 出力OKのものだけ出力する
		result_staff_list = Array.new
		staff_list.each{|data|
			if ( data[:is_output] == true )
				result_staff_list.push(data)
			end
		}

		puts "excel count = #{result_staff_list.size()}"

		Excel.runDuring(false, false) do |excel|

			# 社員数だけ
			result_staff_list.each{|data|

				# 出力先のパスを取得
				out_path = getOutputPath( "#{data[:id]}", "#{data[:abbrev_name]}", "#{data[:create_calendar]}" )

				# テンプレートのブックをコピー
				fsoCopyFile( "#{TEMPLATE_FILE_NAME}", out_path)

				# コピーしたブックを開いて、パラメータの設定
				wb = Excel.openWb(excel, out_path)
				setWsParam( wb, data )

				# 左上にスクロールしておく
				Excel.setScrollWithActiveWindow( excel, 1, 1 )

				# セーブして閉じる
				Excel.saveAndClose( wb )

				# ログ用
				puts "create excel => #{File::basename( out_path )}"
			}
		end
	end

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
		out_path	= File.expand_path( "#{OUT_ROOT}/#{file_name}" )

		return out_path;
	end

	#----------------------------------------------
	# @biref	「報告締切日」を算出
	# @parm		defult_str		デフォルトの文言
	# @parm		dead_line_str	置換する文言
	#----------------------------------------------
	def calcDeadLineString( defult_str, dead_line_str )

		# デフォルトの文字列の 「XXXX」 をパラメータで置換する
		return defult_str.gsub( DEAD_LINE_STR_PREFIX, dead_line_str )
	end

	#----------------------------------------------
	# @biref	パラメータの設定
	# @parm		wb			設定を行うワークブック
	# @parm		param_hash	パラメータを格納したハッシュ
	#----------------------------------------------
	def setWsParam( wb, param_hash )

		return if( Excel.existSheet( wb, SHEET_NAME_TEMPLATE_DATA) == false )

		# 社員名（フルネーム）
		ws_staff	  = wb.worksheets(SHEET_NAME_TEMPLATE_DATA)
		ws_staff.name = "#{param_hash[:abbrev_name]}"

		Excel.setProtectSheet(ws_staff, false)
			Excel.setCellValue( ws_staff, 6, 3, "%03d" % "#{param_hash[:id]}".to_i ) # 数値を3桁に変換
			Excel.setCellValue( ws_staff, 7, 3, "#{param_hash[:name]}" )
			Excel.setCellValue( ws_staff, 8, 3, "#{param_hash[:job_type]}" )
			Excel.setCellValue( ws_staff, 6, 7, "#{param_hash[:gender]}" )
			Excel.setCellValue( ws_staff, 7, 7, "#{param_hash[:age]}" )
			Excel.setCellValue( ws_staff, 8, 7, "#{param_hash[:joining_time]}" )
			Excel.setCellValue( ws_staff, 7, 10, "#{param_hash[:last_month_over_time]}" )
			Excel.setCellValue( ws_staff, 8, 10, "#{param_hash[:last_month_over_time2]}" )
			dead_line_string = calcDeadLineString( Excel.getCellValue(ws_staff, 73, 1), "#{param_hash[:report_dead_line]}" )
			Excel.setCellValue( ws_staff, 73, 1, dead_line_string )
			Excel.setStringColor( ws_staff, 73, 1, dead_line_string, "#{param_hash[:report_dead_line]}" )
		Excel.setProtectSheet(ws_staff, true)
	end
end
