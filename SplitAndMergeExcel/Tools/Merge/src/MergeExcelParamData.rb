# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require "fileutils"
require File.dirname(__FILE__) + "/../../lib/excel.rb"

# ==========================="
# src
# ==========================="
class MergeExcelParamData
	private
	SRC_ROOT 					= File.dirname(__FILE__) + "/.."
	PARAMETER_FILE_PATH	= "#{SRC_ROOT}/MergeParam.xls"

	def initialize()
		@param_is_protected			= "シートを保護するか"
		@param_is_delete_ws_check	= "[届け出チェックシート]を削除するか"
		@param_list					= Array.new
		@param_list.clear
	end

	def setClumn( ws )
		@clumn_is_protected			= Excel.getColumn(ws, @param_is_protected)
		@clumn_is_delete_ws_check	= Excel.getColumn(ws, @param_is_delete_ws_check)
	end

	public
	def getParamList()
		return @param_list
	end

	def setData()
	
		if( File.exist?( "#{PARAMETER_FILE_PATH}" ) == false )
			return
		end
		
		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param = excel.workbooks.open({'filename'=> "#{PARAMETER_FILE_PATH}", 'updatelinks'=> 0})
			ws_param = wb_param.worksheets(1)
	
			# 列番号の設定
			setClumn( ws_param )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do 
			
				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				# パラメータを取得して push
				param = Hash.new
				param[:is_protected]		= Excel.getCellValue(ws_param, recode.row, "#{@clumn_is_protected}".to_i)
				param[:is_delete_ws_check]	= Excel.getCellValue(ws_param, recode.row, "#{@clumn_is_delete_ws_check}".to_i)
				@param_list.push( param )
			end
			
			wb_param.close(0)
		end
	end
	
end