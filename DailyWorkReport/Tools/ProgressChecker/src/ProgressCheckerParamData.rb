# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/excel.rb"

# ==========================="
# src
# ==========================="
class ProgressCheckerParamData
	SRC_ROOT 			= File.dirname(__FILE__) + "/.."
	PARAMETER_FILE_NAME	= "#{SRC_ROOT}/Param.xls"
	PARAM_001			= "チェックする人"
	PARAM_002			= "祝日・休日チェック"
	
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
			ws_param = wb_param.worksheets( "パラメータ" )
	
			# 列番号の設定
			setClumn( ws_param )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do 
			
				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)
				no = Excel.getCellValue(ws_param, recode.row, 1 )
				next if ( no == "" or no == nil)
				
				# パラメータを取得してpush
				custodian			= Excel.getCellValue(ws_param, recode.row, @clumn_custodian )
				is_holiday_check	= Excel.getCellValue(ws_param, recode.row, @clumn_is_holiday_check )

				param = Hash.new
				param[:custodian]		 = custodian
				param[:is_holiday_check] = is_holiday_check
				
				if( !errorCheck( recode.row, param ) )
					@param_list.push( param )
				end

				# 1行しか処理しません
				break
			end
			wb_param.close(0)
		end
	end

	private
	def setClumn( ws )
		@clumn_custodian		= Excel.getColumn(ws, PARAM_001 )
		@clumn_is_holiday_check	= Excel.getColumn(ws, PARAM_002 )
	end
	
	def errorCheck( row_num, param )
	
		is_enable_param = [ true, true ]
		if( param[:custodian] == "" or param[:custodian] == nil )
			is_enable_param[ 0 ] = false
		elsif( param[:is_holiday_check] == "" or param[:is_holiday_check] == nil )
			is_enable_param[ 1 ] = false
		end
		
		if( is_enable_param[ 0 ] && is_enable_param[ 1 ] )
			return false
		end

		error_str = "Parameter Error!!" + "\n"
		
		if( is_enable_param[ 0 ] == false )
			error_str = error_str + "#{row_num}行目 => 「#{PARAM_001}」が未入力です。"
		elsif( is_enable_param[ 1 ] == false )
			error_str = error_str + "#{row_num}行目 => 「#{PARAM_002}」が未入力です。"		
		end
		assertLogPrintFalse( "#{error_str}" )		
		return true
	end	
end