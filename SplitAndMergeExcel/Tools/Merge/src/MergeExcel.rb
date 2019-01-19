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
	IN_ROOT                             = File.dirname(__FILE__) + "/../in"
	OUT_ROOT                          = File.dirname(__FILE__) + "/../out"
	REMAINING_DAYS_WB_NAME  = "残日数管理".encode( Encoding::Windows_31J )

	public
	def initialize()
		@dir_path_list = Array.new
		@file_list_array = Array.new
	end

	def getDirPathList()
		return @dir_path_list
	end

	def main()

		# ファイルが存在していた場合はファイルを削除
		allClearFile( OUT_ROOT, SEARCH_PAT_ARRAY )

		# フォルダ / ファイルパスをリストに設定
		setFileList( @dir_path_list, @file_list_array )

		# excelの処理
		Excel.runDuring(false, false) do |excel|

			# in フォルダにある dir の数だけ処理
			@file_list_array.each{ |file_list|

				puts "\ndirectory => #{File.basename(File.dirname(file_list[0]))}"
				puts "-------------------------------------------------"

				# 新規ワークブックを作成・必要なシートをコピー
				wb_merge = excel.workbooks.add()

				# 最初に作成されたワークシート名を変更しておく
				(1..3).each{|num|
					wb_merge.worksheets("Sheet#{num}").name = "#{num}#{num}#{num}#{num}#{num}"
				}

				# dir にあるエクセルの数だけ処理
				file_list.each_with_index{ |file_path, file_count|

					# excel を開いてワークシートをコピー
					wb_src = excel.workbooks.open({'filename'=> "#{file_path}", 'updatelinks'=> 0})
					Excel.sheetCopyNumber( wb_src, 1, wb_merge, file_count+1 )
					wb_src.close()
					puts "merge => " + "#{File.basename(file_path)}"
				}

				# 最初に作成されたワークシートを削除
				(1..3).each{|num|
					wb_merge.worksheets("#{num}#{num}#{num}#{num}#{num}").delete()
				}

				# 最初のシートをアクティブにして終了
				wb_merge.worksheets(1).Activate

				dir_path	= File.dirname( file_list[0] )
				file_name	= File.basename( dir_path )

				# ファイル名の拡張子を取得
				ext_name = File.extname( file_list[0] )

				# 保存して閉じる
				out_path = "#{OUT_ROOT}" + "/" + "#{file_name}" + "#{ext_name}"
				puts "output => " + "#{File.basename(out_path)}"
				fso	= WIN32OLE.new('Scripting.FileSystemObject')

				if( ext_name == ".xls" )
					wb_merge.saveAs( "#{fso.GetAbsolutePathName( out_path )}", -4143 )
				else
					wb_merge.saveAs( "#{fso.GetAbsolutePathName( out_path )}" )
				end
				wb_merge.close(0)
			}
		end
	end

	private
	def sortWorksArray( src_array, dst_array )

		# 残日数管理対応
		array_size	= src_array.size()
		(1..array_size+1).each { |count|

			#数字順にするために
			src_array.each{ |path|

				file_name	= File.basename( path )
                wb_name     = REMAINING_DAYS_WB_NAME
				name_array	= file_name.gsub( "#{wb_name}", "" ).split( "_" )
				number		= name_array[0]

				if( count == number.to_i )
					dst_array.push( path )
					break
				end
			}
		}
	end

	def setFileList( dir_list, file_list_array )

		# パスの変換(\\ => /)
		check_dir = IN_ROOT.gsub( "\\", "/" )


		# in フォルダ以下のフォルダをチェック
		Dir.glob( "#{check_dir}/**" ) do |dir_path|

			if ( File::ftype( dir_path ) != "directory" )
				warning_str = "[ #{File.basename(dir_path)} ]" + "\n"
				warning_str += "in フォルダ に フォルダ以外のものが配置されてます!!" + "\n"
				warning_str += "フォルダ 以下にマージしたいファイルを配置して下さい"
				warningLogPrint( "#{warning_str}" )
				next
			end

			dir_list.push( dir_path )

			# パターンにマッチするファイルパスを追加
			file_list_temp = getSearchFileList( dir_path, SEARCH_PAT_ARRAY )

			# 残日数管理エクセル対応
			file_list = Array.new
			if( file_list_temp.size() != 0 )
				sortWorksArray( file_list_temp, file_list )
			end

			# 上記の対応がなされなかった場合は通常処理
			if( file_list.size() <= 0 )
				file_list_temp.each { |path|
					file_list.push( path )
				}
			end

			# ファイルがなかった場合は削除
			if( file_list.size() == 0 )
				dir_list.delete( dir_path )

				if( file_list.size == 0 )
					warning_str = "not found merging file !!" + "\n"
					warning_str += "フォルダ [#{File.basename(dir_path)}] にファイルがあるかお確かめ下さい"
					warningLogPrint( "#{warning_str}" )
				end
			else
				file_list_array.push( file_list )
			end
		end

		# ascii順に並び替え
		dir_list.sort!
	end

	#----------------------------------------------
	# @biref	出力先のパスを取得
	# @parm		file_name	作業月報名
	# @parm		calender	月報日時
	#----------------------------------------------
	def getOutputPath( file_name, calender )
		output_path = "#{OUT_ROOT}/#{file_name}_#{calender}"
		return ( output_path.gsub( ".xlsm", ".xlsx" ) )
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
end

=begin
		# パラメータの適応
		applyParamMergeWb( wb_merge, data )

		# ファイル名から一部を拝借して設定
		file_name_info = @file_path_list[0].split( "_" )
		out_path = getOutputPath( file_name_info[2], file_name_info[3] )

		puts "output merge excel = #{out_path}"
=end
