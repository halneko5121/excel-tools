# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../Template/src/TemplateExcelParamData.rb"
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/TemplateUpdateParamData.rb"
require File.dirname(__FILE__) + "/TemplateUpdate.rb"

# ==========================="
# Const
# ==========================="
TITLE	= "TempleteUpdate"
VER		= "1.0.0"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# パラメータの設定(テンプレート)
		template_param = TemplateExcelParamData.new
		
		# パラメータの設定
		template_update_param = TemplateUpdateParamData.new

		# テンプレートの書式に設定し直し
		template_update	= TemplateUpdate.new
		template_update.update( template_param.getStaffList(), template_update_param.getParamList() )
	}
	
end
