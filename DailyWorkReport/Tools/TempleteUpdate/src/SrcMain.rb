# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../Template/src/ExcelParamData.rb"
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/TemplateUpdate.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "TempleteUpdate"
VER		= "1.0.5"
TAMPLATE_PARAM_FILE_NAME	= File.dirname(__FILE__) + "/../../Template/TemplateParam.xls"
PARAMETER_FILE_NAME			= File.dirname(__FILE__) + "/../TemplateUpdateParam.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータの設定(テンプレート)
		param_hash = { name: "社員名", abbrev_name: "略名", pass: "pass", joining_time: "作成年月" }
		template_param	= ExcelParamData.new(TAMPLATE_PARAM_FILE_NAME, "社員ごとの設定", param_hash)

		# パラメータの設定
		param_hash = { src_range: "旧フォーマットコピーセル範囲", dst_range: "新フォーマットペーストセル範囲" }
		template_update_param = ExcelParamData.new(PARAMETER_FILE_NAME, "パラメータ", param_hash)

		# テンプレートの書式に設定し直し
		template_update	= TemplateUpdate.new
		template_update.update( template_param.getParamList(), template_update_param.getParamList() )
	}

end
