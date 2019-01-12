# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require "fileutils"	
require "find"
require File.dirname(__FILE__) + "/../../lib/excel.rb"

# ==========================="
# src
# ==========================="
class TemplateExcelParamData
	private
	SRC_ROOT 			= File.dirname(__FILE__) + "/.."
	PARAMETER_FILE_NAME	= "#{SRC_ROOT}/TemplateParam.xls"

	def setClumn( ws )
		@clumn_id					= Excel.getColumn(ws, "社員番号")
		@clumn_name					= Excel.getColumn(ws, "社員名")
		@clumn_abbrev_name			= Excel.getColumn(ws, "略名")
		@clumn_job_type				= Excel.getColumn(ws, "職種")
		@clumn_gender				= Excel.getColumn(ws, "性別")
		@clumn_joining_time			= Excel.getColumn(ws, "勤続年数")
		@clumn_age					= Excel.getColumn(ws, "年齢")

		@clumn_last_mouth_over_time	= Excel.getColumn(ws, "先月度時間外労働")
		@clumn_last_mouth_over_time2= Excel.getColumn(ws, "先々月度時間外労働")
		@clumn_create_calendar		= Excel.getColumn(ws, "作成年月")
		@clumn_report_dead_line		= Excel.getColumn(ws, "報告締切日")
		@clumn_is_output			= Excel.getColumn(ws, "出力するか")
	end

	def errorCheck()
	
		@staff_list.each { |staff|

			if( staff[:id] == "" or staff[:id] == nil )
				error_str = "Parameter Error!!"
				error_str = error_str + "「社員番号」が未入力です。"
				assertLogPrintFalse( "#{error_str}" )
			elsif( staff[:abbrev_name] == "" or staff[:abbrev_name] == nil )
				error_str = "Parameter Error!!"
				error_str = error_str + "「略名」が未入力です。"
				assertLogPrintFalse( "#{error_str}" )
			elsif( staff[:create_calendar] == "" or staff[:create_calendar] == nil or staff[:create_calendar] == 0 )
				error_str = "Parameter Error!!"
				error_str = error_str + "「作成年月」が未入力です。"
				assertLogPrintFalse( "#{error_str}" )
			end
		}
	end

	public
	def initialize()
		@staff_list = Array.new
		@staff_list.clear
		
		assertLogPrintNotFoundFile( PARAMETER_FILE_NAME )
	end

	def getStaffList()
		return @staff_list
	end
	
	def setData()
		
		Excel.runDuring(false, false) do |excel|

			# パラメータ用 excel を開く
			wb_param = excel.workbooks.open({'filename'=> "#{PARAMETER_FILE_NAME}", 'updatelinks'=> 0})
			ws_param = wb_param.worksheets("社員ごとの設定")
	
			# 列番号の設定
			setClumn( ws_param )

			# レコードの数だけ
			for recode in ws_param.UsedRange.Rows do 
			
				# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
				next if (recode.row == 1 or recode == "" or recode == nil)

				name = Excel.getCellValue(ws_param, recode.row, "#{@clumn_name}".to_i)			
				next if (name == "" or name == nil)

				is_output = Excel.getCellValue(ws_param, recode.row, "#{@clumn_is_output}".to_i)				
				is_output = is_output.encode( Encoding::UTF_8 ) 
				next if (is_output == "" or is_output == nil or is_output.index( "◯" ) == nil )

				# パラメータを取得してpush
				staff = Hash.new
				staff[:id]						= Excel.getCellValue(ws_param, recode.row, "#{@clumn_id}".to_i).to_i
				staff[:name]					= "#{name}"
				staff[:abbrev_name]				= Excel.getCellValue(ws_param, recode.row, "#{@clumn_abbrev_name}".to_i)
				staff[:job_type]				= Excel.getCellValue(ws_param, recode.row, "#{@clumn_job_type}".to_i)
				staff[:gender]					= Excel.getCellValue(ws_param, recode.row, "#{@clumn_gender}".to_i)
				staff[:joining_time]			= Excel.getCellValue(ws_param, recode.row, "#{@clumn_joining_time}".to_i)
				staff[:age]						= Excel.getCellValue(ws_param, recode.row, "#{@clumn_age}".to_i)
				staff[:last_mouth_over_time]	= Excel.getCellValue(ws_param, recode.row, "#{@clumn_last_mouth_over_time}".to_i).to_s
				staff[:last_mouth_over_time2]	= Excel.getCellValue(ws_param, recode.row, "#{@clumn_last_mouth_over_time2}".to_i).to_s
				staff[:create_calendar]			= Excel.getCellValue(ws_param, recode.row, "#{@clumn_create_calendar}".to_i).to_i
				staff[:report_dead_line]		= Excel.getCellValue(ws_param, recode.row, "#{@clumn_report_dead_line}".to_i)
				@staff_list.push( staff )
			end
			wb_param.close(0)
			
			if ( @staff_list.size() == 0 )
				error_str = "[ TemplateParam.xls ] のパラメータは設定されていますか?\n"
				error_str += "社員ごとの設定が見当たりませんでした"
				assertLogPrintFalse( error_str )
			end
		end
		
		errorCheck()
	end
	
end