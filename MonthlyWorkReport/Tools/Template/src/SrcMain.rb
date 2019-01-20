# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + '/../../lib/AppModule.rb' )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/ExcelParamData.rb' )
require File.expand_path( File.dirname(__FILE__) + '/TemplateExcelCreate.rb' )

# ==========================="
# Const
# ==========================="
TITLE	= "MonthlyWorkTemplete"
VER		= "1.1.1"
PARAMETER_FILE_NAME	= File.dirname(__FILE__) + "/../TemplateParam.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータを取得する(社員ごと)
		param_hash = { id: "社員番号", name: "社員名", abbrev_name: "略名", joining_time: "入社時期", create_calendar: "作成年月", period: "月報期間" }
		template_param = ExcelParamData.new(PARAMETER_FILE_NAME, "社員ごとの設定", param_hash)

		# テンプレートのデータを元に excel の生成
		create_excel = TemplateExcelCreate.new
		create_excel.createExcel( template_param.getParamList() )
	}

end
