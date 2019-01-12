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
TITLE	= "Excel Merge"
VER		= "1.0.4"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータを取得する
		merge_param	= MergeExcelParamData.new
		merge_param.setData()

		param = Array.new
		param = merge_param.getParamList()

		# マージを行う
		merge_script = MergeExcel.new
		merge_script.main( param )
	}
	
end

