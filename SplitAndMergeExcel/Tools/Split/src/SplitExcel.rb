# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require "fileutils"
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/Define.rb"
require File.dirname(__FILE__) + "/CommonParam.rb"
require File.dirname(__FILE__) + "/SplitDefault.rb"
require File.dirname(__FILE__) + "/SplitWorks.rb"

# ===========================
# Const
# ===========================
PARAMETER_FILE_NAME	 = File.expand_path(File.dirname(__FILE__)) + "/../SplitParam.xls"
SPLIT_WORKS_WB_NAME	 = [ "有休*", "代休*", "振休*", "管理表*" ]

# ===========================
# src
# ===========================
class SplitExcel
	public
	def initialize( fso, excel, in_root, out_root )
		@fso		  = fso
		@excel		  = excel
		@in_root	  = in_root
		@out_root	  = out_root
		@common_param = CommonParam.new( excel, PARAMETER_FILE_NAME )
		@split_pattern= @common_param.getParam( "分割パターン" ).to_i
		puts ""
		puts "分割パターン : #{@split_pattern}"
	end

	# --------------------------------------------
	# 分割パターンごとの事前チェック
	# --------------------------------------------
	def splitAllFile( file_list )

		# in フォルダにある excel の数だけ処理
		file_list.each{ |file_path|

			# ファイル名の拡張子を取得
			ext_name = File.extname( file_path )
			
			# ファイル名のフォルダを作成
			dir_name = File.basename( file_path, "#{ext_name}" )			
			out_dir	 = "#{@out_root}/#{dir_name}"
			FileUtils.mkdir_p( "#{out_dir}" )
			
			# エラーチェック
			splitErrorCheck( @split_pattern )

			puts "========== split : #{dir_name} =========="
			
			# 分割したいブックを開いて分割
			src_wb = @excel.workbooks.open({'filename'=> @fso.GetAbsolutePathName( file_path ), 'updatelinks'=> 0})
			split( src_wb, out_dir, ext_name )
			src_wb.close(0)
		}
	end

	private
	# --------------------------------------------
	# 分割パターンごとの事前チェック
	# --------------------------------------------
	def splitErrorCheck( split_pattern )

		case split_pattern
		when 1
			file_list = getSearchFile( "#{@in_root}", SPLIT_WORKS_WB_NAME )
			if( file_list.size() == 0 )
				error_str = "[有給管理表分割]の際は\n"
				error_str += "ファイル名に[振休][管理表]が入ったexcelが必要です"
				assertLogPrintFalse( "#{error_str}" )
			end
		end
	end

	# --------------------------------------------
	# excel　の分割
	# --------------------------------------------
	def split( src_wb, out_dir, ext_name )

		is_ws_protect	= @common_param.getParam( "シートの保護" )
		protect_pass	= @common_param.getParam( "保護パスワード" )
		
		# パラメータ用 excel を開く
		wb_param = @excel.workbooks.open({'filename'=> @fso.GetAbsolutePathName( PARAMETER_FILE_NAME ), 'updatelinks'=> 0})

		# パターンごとに分割
		case @split_pattern
		when 0
			split_default = SplitDefault.new()			
			split_default.split( @excel, @fso, src_wb, out_dir, ext_name, is_ws_protect )
		when 1
			SPLIT_WORKS_WB_NAME.each{ |pattern|
                pattern = pattern.encode( Encoding::Windows_31J )
				if( src_wb.name.match( /#{pattern}/ ) != nil )
					split_work = SplitWork.new()
					split_work.split( @excel, src_wb, out_dir, ext_name, wb_param.worksheets("分割パターン1"), is_ws_protect, protect_pass )
					break
				end			
			}
		else
			error_str = "サポートされていない分割パターンです！ => #{@split_pattern}(0~1)"
			assertLogPrintFalse( "#{error_str}" )
		end

		# Excel が分割されなかった場合はフォルダを削除する
		search_pat_list = getSearchPatternList( out_dir, SEARCH_PAT_ARRAY )
		count = 0
		Dir.glob( search_pat_list ) do |file_path|
			count += 1
		end
		if( count == 0 )
			FileUtils.rm_r( Dir.glob( "#{out_dir}" ) )
		end
			
		wb_param.close(0)
	end
end
