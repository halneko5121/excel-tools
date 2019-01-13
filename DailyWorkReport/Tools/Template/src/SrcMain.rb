# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.expand_path( File.dirname(__FILE__) + '/../../lib/AppModule.rb' )
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )
require File.expand_path( File.dirname(__FILE__) + '/TemplateExcelParamData.rb' )
require File.expand_path( File.dirname(__FILE__) + '/HolidayExcelParamData.rb' )
require File.expand_path( File.dirname(__FILE__) + '/TemplateExcelCreate.rb' )

# ==========================="
# Const
# ==========================="
TITLE	= "DailyWorkTemplete"
VER		= "1.1.0"
PARAMETER_FILE_NAME	= File.dirname(__FILE__) + "/../TemplateParam.xls"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

        # パラメータを取得する(祝日)
		holiday_param		= HolidayExcelParamData.new(PARAMETER_FILE_NAME, "祝日設定", "祝日")
		param_list			= holiday_param.getParamList()

		# パラメータを取得する(社員ごと)
		param_hash = { name: "社員名", abbrev_name: "略名", pass: "pass", joining_time: "作成年月" }
		template_param		= TemplateExcelParamData.new(PARAMETER_FILE_NAME, "社員ごとの設定", param_hash)
		staff_data_list		= template_param.getParamList()

		# テンプレートのデータを元に excel の生成
		create_excel = TemplateExcelCreate.new
		create_excel.createExcel( staff_data_list, param_list )
	}
end
