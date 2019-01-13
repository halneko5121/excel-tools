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

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

        # パラメータを取得する(祝日)
		holiday_param		= HolidayExcelParamData.new
		param_list			= holiday_param.getParamList()

		# パラメータを取得する(社員ごと)
		template_param		= TemplateExcelParamData.new
		staff_list			= template_param.getStaffList()

		# テンプレートのデータを元に excel の生成
		create_excel = TemplateExcelCreate.new
		create_excel.createExcel( staff_list, param_list )
	}
end
