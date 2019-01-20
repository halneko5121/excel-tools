# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.expand_path( File.dirname(__FILE__) + "/../../lib/AppModule.rb" )
require File.expand_path( File.dirname(__FILE__) + "/TemplateUpdate.rb" )

# ==========================="
# Const
# ==========================="
TITLE	= "TempleteUpdate"
VER		= "1.0.4"

# ==========================="
# src
# ==========================="
if ( __FILE__ == $0 )

	AppModule.main( TITLE,  VER ) {

		# テンプレートの書式に設定し直し
		template_update	= TemplateUpdate.new
		template_update.update()
	}

end
