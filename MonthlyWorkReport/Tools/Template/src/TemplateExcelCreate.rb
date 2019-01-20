# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/util.rb' )

# ==========================="
# src
# ==========================="
class TemplateExcelCreate
	private
	OUT_ROOT 					= File.dirname(__FILE__) + "/../../../Users"
	TEMPLATE_FILE_NAME			= File.dirname(__FILE__) + "/../../Template/1-UP作業月報_template.xlsx"
	SHEET_NAME_PRORATED_TABLE	= "区分別按分表"
	NEEDLESS_SHEET_NAMES 		= [ "業務別月報", "未定" ]

	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )
	end

	def createExcel( staff_list )

		# ファイルが存在していた場合はファイルを削除
		pattern = [ "*.xlsx" ]
		allClearFile("#{OUT_ROOT}", pattern)

		puts "excel count = #{staff_list.size()}"

		Excel.runDuring(false, false) do |excel|

			# 社員数だけ
			staff_list.each{|data|

				# 出力先のパスを取得
				out_path = getOutputPath( "#{data[:id]}", "#{data[:abbrev_name]}", "#{data[:create_calendar]}" )

				# テンプレートのブックをコピー
				fsoCopyFile( "#{TEMPLATE_FILE_NAME}", out_path )

				# コピーしたブックを開く
				wb = Excel.openWb( excel, out_path )

				# 不要なシートを削除
				NEEDLESS_SHEET_NAMES.each { |ws_name|
					wb.worksheets( ws_name ).delete()
				}

				# パラメータの設定
				setWsParamStaffSheet( wb, data )
				setWsParamProratedTable( wb, data )

				# シートは非表示にしておく  / [区分別按分表][届書チェック]
				Excel.setVisibleSheet( wb.worksheets("#{SHEET_NAME_PRORATED_TABLE}"), false )

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
		abbrev_name	= abbrev_name.encode( Encoding::UTF_8 )
		file_name	= "#{staff_id}_#{abbrev_name}_1-UP作業月報_#{calendar}.xlsx".encode(Encoding::Windows_31J)
		out_path	= "#{OUT_ROOT}/#{file_name}"
		out_path	=  File.expand_path( out_path )

		return out_path;
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
		Excel.setCellValue( ws_staff, 2, 10, "#{param_hash[:name]}" ) # 氏名

		# 2013xx => [2013][xx]に分割
		str_year_month	= splitYearMonth( "#{param_hash[:create_calendar]}" )
		Excel.setCellValue( ws_staff, 2, 5, str_year_month[0].to_i )
		Excel.setCellValue( ws_staff, 2, 6, str_year_month[1].to_i )
		Excel.setCellValue( ws_staff, 2, 7, "#{param_hash[:period]}" ) # 期間

		# セルをロック（編集不可）にしてシートを保護
		(5..7).each { |index_column|
			Excel.setLockCell(ws_staff, 2, index_column, true)
		}
		Excel.setProtectSheet(ws_staff, true)
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
		Excel.setCellValue( ws_propateed, 5, 22, "#{param_hash[:abbrev_name]}" )
		Excel.setCellValue( ws_propateed, 5, 23, 0 )
		Excel.setCellValue( ws_propateed, 6, 1, Excel.getCellValue( ws_propateed, 5, 1 ) )
		Excel.setCellValue( ws_propateed, 6, 22, "#{param_hash[:abbrev_name]}" )
		(22..23).each { |index_column|
			Excel.setVisibleColumns(ws_propateed, index_column, false)
		}

		# 入社時期
		if( param_hash[:joining_time] != nil )
			Excel.setAddCommentCell( ws_propateed, 5, 1, "#{param_hash[:joining_time]}" )
		end
	end
end
