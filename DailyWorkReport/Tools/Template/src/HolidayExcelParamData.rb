# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )

# ==========================="
# src
# ==========================="
class HolidayExcelParamData
	SRC_ROOT 					= File.dirname(__FILE__) + "/.."
	PARAMETER_FILE_NAME	= "#{SRC_ROOT}/TemplateParam.xls"
	
	public
	def initialize()
		@param_list = Array.new
		@param_list.clear
		assertLogPrintNotFoundFile( PARAMETER_FILE_NAME )
		
		setData()
	end

	def getParamList()
		return @param_list
	end

	def setData()
		
		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param = Excel.openWb( excel, PARAMETER_FILE_NAME )
			ws_param = wb_param.worksheets( "祝日設定" )
	
			# 列番号の設定
			setClumn( ws_param )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do 
			
				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)
                
				# パラメータを取得してpush
                holiday               = Excel.getCellValue(ws_param, recode.row, "#{@clumn_holiday}".to_i )
				next if ( holiday == "" or holiday == nil )
                
                param = Hash.new
                param[:holiday] = holiday
				@param_list.push( param )
			end
			wb_param.close(0)			
		end
		
		errorCheck()
	end

	private
	def setClumn( ws )
		@clumn_holiday	= Excel.getColumn(ws, "祝日")
	end

	def errorCheck()

        if( @param_list.count <= 0 )
            error_str = "Parameter Error!!\n"
            error_str = error_str + "祝日が1つも設定されていません"
            assertLogPrintFalse( "#{error_str}" )
        end
	end
	
end