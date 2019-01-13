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
	def initialize(wb_path)
		@wb_path = wb_path
		@staff_list = Array.new
		@staff_list.clear

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
			ws_param = wb_param.worksheets( "社員ごとの設定" )

			# 列番号の設定
			setClumn( ws_param )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do

				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				name = Excel.getCellValue(ws_param, recode.row, "#{@clumn_name}".to_i)

				next if (name == "" or name == nil)

				# パラメータを取得してpush
				staff = Hash.new
				staff[:name]			= "#{name}"
				staff[:abbrev_name]		= Excel.getCellValue(ws_param, recode.row, "#{@clumn_abbrev_name}".to_i)
				staff[:pass]			= Excel.getCellValue(ws_param, recode.row, "#{@clumn_pass}".to_i )
				staff[:joining_time]	= Excel.getCellValue(ws_param, recode.row, "#{@clumn_joining_time}".to_i)
				@staff_list.push( staff )
			end
			wb_param.close(0)
		end

		errorCheck()
	end

	private
	def setClumn( ws )
		@clumn_name				= Excel.getColumn(ws, "社員名")
		@clumn_abbrev_name	= Excel.getColumn(ws, "略名")
		@clumn_pass				= Excel.getColumn(ws, "pass")
		@clumn_joining_time	= Excel.getColumn(ws, "作成年月")
	end

	def errorCheck()

		@staff_list.each { |staff|

			if( staff[:abbrev_name] == "" or staff[:abbrev_name] == nil )
				error_str = "Parameter Error!!"
				error_str = error_str + "「略名」が未入力です。"
				assertLogPrintFalse( "#{error_str}" )
			elsif( staff[:joining_time] == "" or staff[:joining_time] == nil or staff[:joining_time] == 0 )
				error_str = "Parameter Error!!"
				error_str = error_str + "「作成年月」が未入力です。"
				assertLogPrintFalse( "#{error_str}" )
			end
		}
	end

end
