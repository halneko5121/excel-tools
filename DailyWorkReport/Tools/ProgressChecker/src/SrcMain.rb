# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../Template/src/TemplateExcelParamData.rb"
require File.dirname(__FILE__) + "/../../Template/src/HolidayExcelParamData.rb"
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/ProgressChecker.rb"
require File.dirname(__FILE__) + "/ProgressCheckerParamData.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "ProgressChecker"
VER		= "1.0.0"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータの設定(チェック用)
		check_data		 = ProgressCheckerParamData.new
		custodian		 = check_data.getParamList()[0][:custodian]
		is_holiday_check = check_data.getParamList()[0][:is_holiday_check]

		# パラメータの設定(テンプレート)
		template_param = TemplateExcelParamData.new
		
		# どこまで確認したか。の確認チェック
		checker	= ProgressChecker.new

		# パラメータの設定(祝日)
		if( is_holiday_check )
			holiday_param = HolidayExcelParamData.new	
			checker.exe( template_param.getStaffList(), custodian, holiday_param.getParamList() )
		else
			checker.exe( template_param.getStaffList(), custodian, nil )
		end		
	}
	
end
