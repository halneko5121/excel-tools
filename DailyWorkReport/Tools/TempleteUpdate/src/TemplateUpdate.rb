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
	IN_ROOT 			= File.dirname(__FILE__) + "/../in"
	OUT_ROOT 			= File.dirname(__FILE__) + "/../out"
	TEMPLATE_FILE_PATH	= File.dirname(__FILE__) + "/../../Template/Template.xlsx"
	COPY_RANGE			= "A1:O50"

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

	def update( template_param_list, param_list )

		Excel.runDuring(false, false) do |excel|

			# テンプレートブックを開く
			fso = WIN32OLE.new('Scripting.FileSystemObject')
			wb_templete	= Excel.openWb( excel, TEMPLATE_FILE_PATH )
			ws_templete	= wb_templete.worksheets( SHEET_NAME_TEMPLATE_DATA )

			# ファイルの数だけ
			@file_list.each { |file_path|

				pass = searchPassword( file_path, template_param_list)
				error_path = File.expand_path(file_path.encode( Encoding::UTF_8 ))
				assertLogPrint(pass != nil, "#{error_path} の pass が不明です" )
				src_wb_staff = Excel.openWb( excel, file_path, pass )

				# 元のブックをコピー
				out_path = getOutputPath( file_path )
				fso.CopyFile( file_path, out_path )
				dst_wb_staff = Excel.openWb( excel, out_path, pass )

				# フォーマットを更新
				monthly_days = getMonthlyDays( template_param_list )
				excelFormatUpdate( ws_templete, src_wb_staff, dst_wb_staff, param_list, monthly_days )

				# 更新したものをセーブして閉じる
				src_wb_staff.close()
				dst_wb_staff.save()
				dst_wb_staff.close()

				# 同じファイル名のブックを開く対策。元に戻す
				rename_path = out_path.gsub( "dst_", "" )
				File.rename( out_path, rename_path )

				# ログ用
				puts "update excel => #{File.basename( rename_path )}"
			}
			wb_templete.close()
		end
	end

	private
	#----------------------------------------------
	# @biref	src excel のセルを dst にコピー
	#----------------------------------------------
	def excelFormatUpdate( ws_templete, src_wb_staff, dst_wb_staff, param_list, monthly_days )

		ws_count		= src_wb_staff.worksheets.count
		ws_manual_count	= ws_count - monthly_days

		# シートの数だけ
		(1..ws_count).each{ |num|

			# マニュアルシートの数以上のシート番号になったら
			# @note:マニュアルシートは最初にある前提
			next if( num < ws_manual_count )

			src_ws = src_wb_staff.worksheets( num )
			dst_ws = dst_wb_staff.worksheets( num )

			# 新しいテンプレートに従って、シートを更新
			Excel.rangeCopyFast( ws_templete, COPY_RANGE, dst_ws, COPY_RANGE )

			# 指定範囲のセルをコピペ
			param_list.each { |param|

				# 元の書き込みデータを設定
				src_range = "#{param[:src_range]}"
				dst_range = "#{param[:dst_range]}"
				Excel.rangeCopyFast( src_ws, src_range, dst_ws, dst_range )
			}
		}
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

		return nil
	end

	#----------------------------------------------
	# テンプレートパラメータから日数を返す
	#----------------------------------------------
	def getMonthlyDays( template_param_list )

		monthly_days = 0
		template_param_list.each { |param|

			str_calendar	= splitYearMonth("#{param[:joining_time]}")
			year			= str_calendar[0].to_i
			mouth			= str_calendar[1].to_i

			# 指定月の日数を設定
			monthly_days = getMonthlyDayCount( year, mouth )
			break
		}

		return monthly_days
	end

	#----------------------------------------------
	# 出力パスを取得する
	#----------------------------------------------
	def getOutputPath( src_file_path )

		out_path = src_file_path.gsub( "in/", "out/dst_" )
		out_path = File.expand_path( "#{out_path}" )
		return out_path
	end
end
