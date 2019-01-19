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

		# [in] にある excel のファイルリストを作成
		pattern_array =[ "*.xlsx" ]
		@file_list = getSearchFileList("#{IN_ROOT}", pattern_array)
		if( @file_list.size() == 0 )
			assertLogPrintFalse( "in フォルダにファイルがありません" )
		end
	end

	def execute( template_param_list, custodian )

		Excel.runDuring(false, false) do |excel|

			# ファイルの数だけ
			log_text = File.open( "./log.txt", "w+" )
			@file_list.each { |file_path|

				pass = searchPassword( file_path, template_param_list )
				wb_staff = Excel.openWb( excel, file_path, pass )

				# ログ用
				releasePuts( "check excel => #{File.basename( file_path )}", log_text )
				releasePuts( "----------------------------", log_text )

				# 各シートをチェック
				wb_staff.worksheets.each { |ws|
					checkProgress( ws, custodian, log_text )
				}
				wb_staff.close()
				releasePuts( "", log_text )
			}
			log_text.close

			releasePuts( "結果は [log.txt] にも出力しています" )
			releasePuts( "検索に利用するなりうまくご利用頂ければ" )
		end
	end

	private
	#----------------------------------------------
	# 進捗をチェックする
	#----------------------------------------------
	def checkProgress( ws, custodian, log_text )

		# 「xx月XX日」の形式になってるシートだけチェック
		utf_8_ws_name = ws.name.encode( Encoding::UTF_8 )
		return if( utf_8_ws_name.match( /(\d+月\d+日)/ ) == nil )

		# 祝日
		holidays_str = ""
		if( utf_8_ws_name.include?( "(土)" ) || utf_8_ws_name.include?( "(日)" ) )
			holidays_str = "(休日)"
		elsif( Excel.isWsColorRed( ws ) )
			holidays_str = "(祝日)"
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
end
