# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../Template/src/ExcelParamData.rb"
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/ProgressChecker.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "ProgressChecker"
VER		= "1.0.0"
TAMPLATE_PARAM_FILE_NAME = File.dirname(__FILE__) + "/../../Template/TemplateParam.xls"
PARAM_FILE_NAME			 = File.dirname(__FILE__) + "/../Param.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータの設定(チェック用)
		param_hash		 = { custodian: "チェックする人", is_holiday_check: "祝日・休日チェック" }
		check_data		 = ExcelParamData.new(PARAM_FILE_NAME, "パラメータ", param_hash)
		custodian		 = check_data.getParamList()[0][:custodian]
		is_holiday_check = check_data.getParamList()[0][:is_holiday_check]

		# パラメータの設定(テンプレート)
		template_param_hash = { id: "社員番号", name: "社員名", abbrev_name: "略名", pass: "pass", joining_time: "作成年月" }
		template_param		= ExcelParamData.new(TAMPLATE_PARAM_FILE_NAME, "社員ごとの設定", template_param_hash)

		# どこまで確認したか。の確認チェック
		checker	= ProgressChecker.new

		# パラメータの設定(祝日)
		if( is_holiday_check )
			holiday_param_hash	= { holiday: "祝日" }
			holiday_param		= ExcelParamData.new(TAMPLATE_PARAM_FILE_NAME, "祝日設定", holiday_param_hash)
			checker.execute( template_param.getParamList(), custodian, holiday_param.getParamList() )
		else
			checker.execute( template_param.getParamList(), custodian, nil )
		end
	}

end
