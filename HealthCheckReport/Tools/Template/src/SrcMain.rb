# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/../../lib/ExcelParamData.rb"
require File.dirname(__FILE__) + "/TemplateExcelCreate.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "HealthCheckReport"
VER		= "1.0.0"
PARAMETER_FILE_NAME	= File.dirname(__FILE__) + "/../TemplateParam.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータを取得する
		param_hash = {
			id: "社員番号",
			name: "社員名",
			abbrev_name: "略名",
			job_type: "職種",
			gender: "性別",
			joining_time: "勤続年数",
			age: "年齢",
			last_mouth_over_time: "先月度時間外労働",
			last_mouth_over_time2: "先々月度時間外労働",
			create_calendar: "作成年月",
			report_dead_line: "報告締切日",
			is_output: "出力するか"
		}
		template_param	= ExcelParamData.new(PARAMETER_FILE_NAME, "社員ごとの設定", param_hash)

		# テンプレートのデータを元に excel の生成
		create_excel = TemplateExcelCreate.new
		create_excel.createExcel( template_param.getParamList() )
	}

end
