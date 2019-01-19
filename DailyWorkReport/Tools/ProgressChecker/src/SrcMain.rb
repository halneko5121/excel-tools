# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/ExcelParamData.rb"
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/ProgressChecker.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "ProgressChecker"
VER		= "1.0.5"
TAMPLATE_PARAM_FILE_NAME = File.dirname(__FILE__) + "/../../Template/TemplateParam.xls"
PARAM_FILE_NAME			 = File.dirname(__FILE__) + "/../Param.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータの設定(テンプレート)
		template_param_hash = { id: "社員番号", name: "社員名", abbrev_name: "略名", pass: "pass", joining_time: "作成年月" }
		template_param		= ExcelParamData.new(TAMPLATE_PARAM_FILE_NAME, "社員ごとの設定", template_param_hash)

		# どこまで確認したか。の確認チェック
		param_hash		= { custodian: "チェックする人" }
		check_data		= ExcelParamData.new(PARAM_FILE_NAME, "パラメータ", param_hash)
		custodian		= check_data.getParamList()[0][:custodian]
		checker			= ProgressChecker.new
		checker.execute( template_param.getParamList(), custodian )
	}

end
