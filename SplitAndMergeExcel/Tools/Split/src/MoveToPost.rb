# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require "FileUtils"
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/Define.rb"
require File.dirname(__FILE__) + "/SplitWorks.rb"

# ==========================="
# const
# ==========================="
PARAMETER_FILE_NAME	= File.dirname(__FILE__) + "/../../Split/SplitParam.xls"
CONNECT_DRIVE = "Z:"
CONNECT_PATH  = "\\\\mfileserver\\post"

# ==========================="
# class
# ==========================="
class MoveToPost
	public
	def initialize( out_root )
		@moveFilePathHash = {}
		@out_root	= out_root
		@file_list	= ganarateFileList()
		@staff_list = getStaffParam()
	end
	
	def connect( password, user_name )
	
		# 接続する前に確認する
		puts "#{CONNECT_PATH} にログインします"
		puts "よろしいですか?(既にログイン済みの場合はn) [yes -> y][no -> n]\n"
		result = STDIN.gets.chomp

		if( "#{result}" == "y" )
			command = "@NET USE #{CONNECT_DRIVE} #{CONNECT_PATH} #{password} /USER:LO\\#{user_name} /PERSISTENT:NO"
			puts "connect to ... [#{CONNECT_PATH}]"
			system( "#{command}" )
		end
	end

	def printMoveFileList()

		puts ""
		puts "以下、移動先一覧"
		puts "-----------------------------------------------------------------"	
		success_count = 0		
		# ファイル名にスタッフの名前が入っているかチェック
		@file_list.each { |file_path|
		
			# スタッフの名前を検索
			@staff_list.each { |staff|

				staff_name = "#{staff[:check_ws_name]}".encode( Encoding::UTF_8 ) 
				if( file_path.index( "#{staff_name}" ) )

					# 移動先のフォルダを作成
					move_path = CONNECT_PATH + "/#{staff[:post_dir_name]}"
					move_path = move_path.gsub( "\\", "/"  )
					
					if( File.exists?( "#{move_path}" ) == false )
						puts "#{move_path} => 指定のフォルダが見当たりません. スキップします"
						next
					end

					# ファイル名以前のパスをリネーム
					dir_path	= File.dirname( file_path )
					rename_path	= file_path.gsub( "#{dir_path}", "#{move_path}" )				

					if( File.exists?( "#{rename_path}" ) == true )
						puts "#{rename_path} => ファイルは既に存在しています！！"
						next
					end
					success_count += 1
					
					# 情報を設定しておく
					@moveFilePathHash[ "#{staff_name}" ] = "#{file_path},#{rename_path}"
					puts "move: #{File.basename(file_path)} => #{rename_path}"
				end
			}
		}
		
		# post フォルダに移動できないファイルがいくつかあった場合の出力
		printErrorDiffCount( success_count )
		puts "-----------------------------------------------------------------"
	end

	def moveFile()
	
		if( @moveFilePathHash.length == 0 )
			str = "移動可能なファイルが1つもありません\n処理を中断します"
			assertLogPrintFalse( str )
		else

			# 移動する前に確認を取る
			puts "上記の場所にファイルを移動してもよろしいですか？ [yes -> y][no -> n]"
			result = STDIN.gets.chomp

			if( "#{result}" == "y" )

				# 各スタッフごとにファイルを移動
				@moveFilePathHash.each_value {|value|
				
					path_array	= value.split( "," ) # [,] で分割
					src_path	= path_array[0]
					dst_path	= path_array[1]
					FileUtils.move( "#{src_path}", "#{dst_path}" )
					puts "move: #{File.basename(src_path)} => #{dst_path}"
				}
			else
				puts "処理を中断します"
			end
		end
	end

	private
	def printErrorDiffCount( success_count )

		if( @staff_list.size() == success_count )
			return
		end
		
		setConsoleColor( "olive", "white" )
		puts "-----------------------------------------------------------------"	
		puts "移動できないファイルがいくつかありました"
		puts "SplitParam.xls の「postフォルダ名」が合っているか"
		puts "同じ名前のファイルがないか。確認下さい"
		puts "#{CONNECT_PATH} のフォルダ一覧は以下"
		puts "-----------------------------------------------------------------"	
		dst_dir_path = "#{CONNECT_PATH}".gsub( "\\", "/"  )
		Dir.foreach( "#{dst_dir_path}" ) { |f|
		
			# カレントと親は表示しない
			next if (f =~ /^\.{1,2}/ )
			puts "#{f}"			
		}	
	end
	
	def ganarateFileList()

		file_list = Array.new()
		file_list.clear()

		# [out] にある excel のファイルリストを作成
		Dir.glob( "#{@out_root}" + "/有休_代休_振休管理表*" + "/**/" + "*.*" ) do |file_path|
			file_list.push( file_path )
		end
		
		if( file_list.size() == 0 )
			str = "[有休_代休_振休管理表]フォルダがないか。excelがありません"
			assertLogPrintFalse( "#{str}" )
		end
		
		return file_list
	end

	def getStaffParam()
		ws_list = Array.new

		# パラメータを取得
		Excel.runDuring(false, false) { |excel|
		
			split_work = SplitWork.new()
		
			wb_param = Excel.openWb( excel, PARAMETER_FILE_NAME )
			ws_split_paterrn = wb_param.worksheets("分割パターン1")
			split_work.setData( ws_split_paterrn )
			ws_split_paterrn.ole_free
			wb_param.close(0)
			ws_list = split_work.getWorkSheetList()
		}
		return ws_list
	end
	
end