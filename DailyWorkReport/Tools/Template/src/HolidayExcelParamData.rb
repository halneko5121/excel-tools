# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )

# ==========================="
# src
# ==========================="
class HolidayExcelParamData
	public
	def initialize(wb_path, ws_name, param_name)
		@wb_path = wb_path
		@ws_name = ws_name
		@param_name = param_name
		@param_list = Array.new
		@param_list.clear

		assertLogPrintNotFoundFile( @wb_path )
		setData()
	end

	def getParamList()
		return @param_list
	end

	private
	def setData()

		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param 		= Excel.openWb( excel, @wb_path )
			ws_param		= wb_param.worksheets( @ws_name )
			param_column	= Excel.getColumn(ws_param, @param_name)

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do

				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				# パラメータを取得してpush
                cell_param = Excel.getCellValue(ws_param, recode.row, "#{param_column}".to_i )
				next if ( cell_param == "" or cell_param == nil )

                param = Hash.new
                param[:holiday] = cell_param
				@param_list.push( param )
			end
			wb_param.close(0)
		end

		errorCheck()
	end

	def errorCheck()

        if( @param_list.count <= 0 )
            error_str = "Parameter Error!!\n"
            error_str = error_str + "祝日が1つも設定されていません"
            assertLogPrintFalse( "#{error_str}" )
        end
	end

end
