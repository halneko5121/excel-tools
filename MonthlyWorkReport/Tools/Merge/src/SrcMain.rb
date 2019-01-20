# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + "/../../lib/AppModule.rb" )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/ExcelParamData.rb' )
require File.expand_path( File.dirname(__FILE__) + "/MergeExcel.rb" )

# ==========================="
# Const
# ==========================="
TITLE	= "Excel Merge"
VER		= "1.0.5"
PARAMETER_FILE_NAME	= File.dirname(__FILE__) + "/../MergeParam.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータを取得する
		param_hash = { is_protected: "シートを保護するか", is_delete_ws_check: "[届け出チェックシート]を削除するか" }
		merge_param = ExcelParamData.new(PARAMETER_FILE_NAME, "Sheet1", param_hash)

		# マージを行う
		merge_script = MergeExcel.new
		merge_script.main( merge_param.getParamList() )
	}

end
