# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"

# ==========================="
# src
# ==========================="
class ProgressChecker
	private
	IN_ROOT 			= File.dirname(__FILE__) + "/../in"
	TEMPLATE_FILE_PATH	= File.dirname(__FILE__) + "/../../Template/Template.xlsx"

	public
	def initialize()

		assertLogPrintNotFoundFile( TEMPLATE_FILE_PATH )

		@file_list = Array.new()
		# [in] にある excel のファイルリストを作成
		Dir.glob( "#{IN_ROOT}" + "/**/" + "*.xlsx" ) do |file_path|
			@file_list.push( file_path )
		end

		if( @file_list.size() == 0 )
			assertLogPrintFalse( "in フォルダにファイルがありません" )
		end
	end

	def execute( template_param_list, custodian, holiday_param_list )

		Excel.runDuring(false, false) do |excel|

			# ファイルの数だけ
			log_text = File.open( "./log.txt", "w+" )
			@file_list.each { |file_path|

				pass = searchPassword( file_path, template_param_list )
				wb_staff = Excel.openWb( excel, file_path, pass )

				# このブックに含まれてる祝日を抜き出す
				holidays = getIncludeHolyday( wb_staff, holiday_param_list )

				# ログ用
				releasePuts( "check excel => #{File.basename( file_path )}", log_text )
				releasePuts( "----------------------------", log_text )

				# 各シートをチェック
				wb_staff.worksheets.each { |ws|
					checkProgress( ws, custodian, holidays, log_text )
				}
				wb_staff.close()
				releasePuts( "", log_text )
			}
			log_text.close

			releasePuts( "結果は [log.txt]　にも出力しています" )
			releasePuts( "検索に利用するなりうまくご利用頂ければ" )
		end
	end

	private
	#----------------------------------------------
	# 進捗をチェックする
	#----------------------------------------------
	def checkProgress( ws, custodian, holidays, log_text )

		# 「xx月XX日」の形式になってるシートだけチェック
		utf_8_ws_name = ws.name.encode( "UTF-8" )
		return if( utf_8_ws_name.match( /(\d+月\d+日)/ ) == nil )

		# 祝日
		holidays_str = ""
		if( holidays != nil )
			if( isPublicHoiyday( utf_8_ws_name, holidays ) )
				holidays_str = "(祝日)"
			elsif( utf_8_ws_name.include?( "(土)" ) || utf_8_ws_name.include?( "(日)" ) )
				holidays_str = "(休日)"
			end
		end

		is_checked = isCheckedProgress( ws, custodian )
		log_str = ""
		if( is_checked )
			log_str = "#{utf_8_ws_name} => チェック済み#{holidays_str}"
		else
			log_str = "#{utf_8_ws_name} => 未チェック#{holidays_str}"
		end
		releasePuts( log_str, log_text )
	end

	#----------------------------------------------
	# チェック済みかどうか
	# @param	ws			ワークシート
	# @param	custodian	管理人名
	#----------------------------------------------
	def isCheckedProgress( ws, custodian )

		check_cell = Excel.getCellValue( ws, 19, 4)
		is_checked = false
		if( check_cell != nil )

			# 「上司コメント」のひとつ下の行に指定した値があるか？
			value = check_cell.to_s
			if( value.include?( "#{custodian}" ) )
				is_checked = true
			end
		end

		return is_checked
	end

	#----------------------------------------------
	# パスワードを検索する
	#----------------------------------------------
	def searchPassword( file_path, template_param_list )

		template_param_list.each { |param|

			# テンプレートパラメータと一致する excel である
			if( file_path.include?( "#{param[:abbrev_name]}" ) )

				# パスワードが設定されている
				pass = "#{param[:pass]}"
				if( ( pass == nil or pass == "" ) == false )
					return pass
				end
			end
		}

		return ""
	end

	#----------------------------------------------
	# 祝日か
	#----------------------------------------------
	def isPublicHoiyday( ws_name, holidays )

		holidays.each { |holiday|

			if( ws_name.include?( holiday ) )
				return true
			end
		}
		return false
	end

	#----------------------------------------------
	# 指定ブックに含まれる祝日を返す
	#----------------------------------------------
	def getIncludeHolyday( wb, holiday_param_list )

		if( holiday_param_list == nil )
			return nil
		end

		holidays = Array.new()
		holiday_param_list.each { |holiday_param|

			# 「西暦」以外の部分を抜き出す
			split_holiday	= holiday_param[:holiday].split( "/" )
			month_day		= "#{split_holiday[1]}月#{split_holiday[2]}日"

			wb.worksheets.each { |ws|
				utf_8_ws_name = ws.name.encode( "UTF-8" )
				if( utf_8_ws_name.include?( month_day ) )
					holidays.push( month_day )
					break
				end
			}
		}
		return holidays
	end

end
