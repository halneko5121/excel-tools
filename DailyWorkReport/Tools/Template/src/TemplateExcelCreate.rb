# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require "date"
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/util.rb' )

# ==========================="
# src
# ==========================="
class TemplateExcelCreate
	private
	OUT_ROOT 			= File.dirname(__FILE__) + "/../../../Users"
	TEMPLATE_FILE_NAME	= File.dirname(__FILE__) + "/../Template.xlsx"

	public
	def initialize()
		assertLogPrintNotFoundFile( TEMPLATE_FILE_NAME )
		@is_leap_year = false
	end

	def createExcel( staff_list, holiday_list )

		# ファイルが存在していた場合はファイルを削除
		Dir.glob( "#{OUT_ROOT}" + "/**/" + "*.*" ) do |file_path|
			File.delete "#{file_path}"
		end

		puts "excel count = #{staff_list.size()}"

		fso = WIN32OLE.new('Scripting.FileSystemObject')
		Excel.runDuring(false, false) do |excel|

			# 社員数だけ
			staff_list.each_with_index{ |data, index|

				# 出力先のパスを取得
				out_path = getOutputPath( index+1, "#{data[:abbrev_name]}" )

				# テンプレートのブックをコピー
				fso.CopyFile( TEMPLATE_FILE_NAME, out_path )

				# コピーしたブックを開く
				wb = Excel.openWb( excel, out_path )

				# パラメータの設定
				setWsParamStaffSheet( wb, data, holiday_list )

				# パスワードが設定されていたら設定する
				pass = "#{data[:pass]}"
				if( ( pass == nil or pass == "" ) == false )
					wb.password = pass
				end
				wb.save()
				wb.close()

				# ログ用
				puts "create excel => #{File::basename( out_path )}"
			}
		end

		if( @is_leap_year )
			puts ""
			puts "閏年の2月です。閏年設定がされました"
		end
	end

	private
	#----------------------------------------------
	# @biref	出力先のパスを取得
	# @parm		fso				ファイルシステムオブジェクト
	# @parm		number			データ番号
	# @parm		abbrev_name		社員略称
	#----------------------------------------------
	def getOutputPath( number, abbrev_name )

		# 数値を3桁に変換
		staff_num	= "%03d" % number
        abbrev_name	= abbrev_name.encode( Encoding::UTF_8 )
		file_name   = "#{staff_num}_#{abbrev_name}_1-UP作業日報.xlsx".encode(Encoding::Windows_31J)
		out_path	= "#{OUT_ROOT}/#{file_name}"
		out_path	=  File.expand_path( out_path )
		out_path	= out_path.gsub( "\\", "/" )

		return out_path;
	end

	#----------------------------------------------
	# @biref	パラメータの設定(各社員シート)
	# @parm		wb				設定を行うワークブック
	# @parm		param_hash      パラメータを格納したハッシュ
	# @parm		holiday_list    祝日リスト
	#----------------------------------------------
	def setWsParamStaffSheet( wb, param_hash, holiday_list )

		if( Excel.existSheet( wb, "日報" ) == false )
			assertLogPrintFalse( "「日報」シートが存在しません。日報用のシート名を「日報」にして下さい" )
			return
		end

		# 2013xx => [2013][xx]に分割
		str_calendar	= splitYearMonth("#{param_hash[:joining_time]}")
		year			= str_calendar[0].to_i
		month			= str_calendar[1].to_i

		# 閏年かどうか
		if( isLeapYear( year, month ) )
			@is_leap_year = true
		end

		# 「日報」 シート以外のシートの数
		manual_count = wb.worksheets.count() -1

		# 指定月の日付分シートを作成
		monthly_days = getMonthlyDayCount( year, month )
		( 1.. monthly_days ).each { |day|

			new_sheet_index = manual_count + day
			if( day != 1 )
				copy_sheet_index = new_sheet_index - 1
				Excel.sheetCopyNumber( wb, manual_count + 1, wb, copy_sheet_index )
			end

			# シート名設定
			w_day		= calcWeekDay( year, month, day )
			sheet_name	= "#{month}月#{day}日(#{w_day})"
			ws = wb.worksheets( new_sheet_index )
			ws.name = sheet_name

			# シート色設定
			Excel.setSheetColorWithWeekend(ws, w_day)

			# 平日なら祝日チェック
			if isWeekday( year, month, day )
				holiday_list.each { |holiday|
                    if( "#{year}/#{month}/#{day}" == "#{holiday[:holiday]}" )
                        wb.worksheets( new_sheet_index ).Tab.ColorIndex = 3
                        break
                    end
                }
            end
		}

		# 最初のシートをアクティブに
		wb.worksheets( 1 ).Activate
	end
end
