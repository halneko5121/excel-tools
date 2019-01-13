# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="

# ==========================="
# class
# ==========================="
class CommonParam

	def initialize( excel, file_path )

		@paramHash = Hash.new

		# パラメータ用 excel を開く
		fso = WIN32OLE.new('Scripting.FileSystemObject')
		wb_param = excel.workbooks.open({'filename'=> fso.GetAbsolutePathName( "#{file_path}" ), 'updatelinks'=> 0})
		ws_param = wb_param.worksheets("全体設定")

		# レコードの数だけ
		for recode in ws_param.UsedRange.Rows do

			# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
			next if (recode.row == 1 or recode == "" or recode == nil)

			@paramHash[ "分割パターン" ] = Excel.getParamValue(ws_param, "分割パターン", recode.row)
			@paramHash[ "シートの保護" ] = Excel.getParamValue(ws_param, "シートの保護", recode.row)
			@paramHash[ "保護パスワード" ] = Excel.getParamValue(ws_param, "保護パスワード", recode.row)
		end

		wb_param.close(0)
	end

	def getParam( param_name )
		@paramHash[ "#{param_name}" ]
	end
end