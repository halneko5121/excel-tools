# -*- coding: utf-8 -*-

# ===========================
# require
# ===========================
require File.expand_path( File.dirname(__FILE__) + '/../../lib/excel.rb' )

# ==========================="
# src
# ==========================="
class TemplateExcelParamData
	public
	def initialize(wb_path, ws_name, param_name_hash)
		@wb_path = wb_path
		@ws_name = ws_name
		@staff_list = Array.new
		@staff_list.clear

		# パラメータ名を保持
		@param_name_hash = param_name_hash
		assertLogPrintNotFoundFile( @wb_path )
		setData()
	end

	def getStaffList()
		return @staff_list
	end

	def setData()

		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param = Excel.openWb( excel, @wb_path )
			ws_param = wb_param.worksheets( @ws_name )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do

				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				# パラメータを取得してpush
				staff = Hash.new
				@param_name_hash.each  { |key, value|
					column_name = Excel.getColumn(ws_param, "#{value}")
					staff[ :"#{key}" ] = Excel.getCellValue(ws_param, recode.row, "#{column_name}".to_i)
				}
				@staff_list.push( staff )
			end
			wb_param.close(0)
		end
		errorCheck()
	end

	private
	def errorCheck()
		@staff_list.each { |staff|
			@param_name_hash.each  { |key, value|
				data = staff[ :"#{key}" ]
				if( data == "" or data == nil )
					error_str = "Parameter Error!!"
					error_str = error_str + "「#{value}」が未入力です。"
					assertLogPrintFalse( "#{error_str}" )
				end
			}
		}
	end

end
