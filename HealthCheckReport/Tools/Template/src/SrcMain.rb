# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/TemplateExcelParamData.rb"
require File.dirname(__FILE__) + "/TemplateExcelCreate.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "HealthCheckReport"
VER		= "1.0.0"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {
	
		staff_list	= Array.new

		# パラメータを取得する
		template_param	= TemplateExcelParamData.new
		template_param.setData()
		staff_list		= template_param.getStaffList()

		# テンプレートのデータを元に excel の生成
		create_excel = TemplateExcelCreate.new
		create_excel.createExcel( staff_list )
	}

end