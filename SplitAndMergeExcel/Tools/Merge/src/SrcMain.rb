# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/MergeExcelParamData.rb"
require File.dirname(__FILE__) + "/MergeExcel.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "ExcelMerge"
VER		= "1.0.1"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータを取得する
		merge_param	= MergeExcelParamData.new
		merge_param.setData()

	#	param = Array.new
	#	param = merge_param.getParamList()

		# マージを行う
		merge_script = MergeExcel.new
		merge_script.main()

		dir_path_list = merge_script.getDirPathList()
		if( dir_path_list.size() == 0 )
			assertLogPrintFalse( "マージする excel がありませんでした" )
		end
	}

end
