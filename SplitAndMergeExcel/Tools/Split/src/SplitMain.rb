# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/SplitExcel.rb"
require File.dirname(__FILE__) + "/Define.rb"

# ===========================
# Const
# ===========================
TITLE	 = "SplitExcel"
VER		 = "1.0.3"
IN_ROOT	 = File.dirname(__FILE__) + "/../in"
OUT_ROOT = File.dirname(__FILE__) + "/../out"

# ===========================
# src
# ===========================
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# ファイルが存在していた場合はファイルを削除
		allRomoveFile( OUT_ROOT, SEARCH_PAT_ARRAY )

		# パターンにマッチするファイルパスを追加
		file_list = getSearchFile( IN_ROOT, SEARCH_PAT_ARRAY )
		file_list.each { |file_path|
			puts "search => " + "#{File.basename(file_path)}"
		}

		if ( file_list.size() == 0 )
			assertLogPrintFalse( "in フォルダに excel ファイルが見当たりません" )
		end

		# そこにシートを分割したブックを放り込む
		Excel.runDuring(false, false) do |excel|

			# 「全体設定」パラメータの設定
			fso = WIN32OLE.new('Scripting.FileSystemObject')
			split_excel = SplitExcel.new( fso, excel, IN_ROOT, OUT_ROOT )
			split_excel.splitAllFile( file_list )
		end
	}

end
