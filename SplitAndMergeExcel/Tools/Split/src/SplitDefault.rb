# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="

# ==========================="
# class
# ==========================="
class SplitDefault
	public	
	def initialize()
	end
	
	def split( excel, fso, src_wb, out_dir, ext_name, is_ws_protect )

		# ワークシートの数だけブックを作成する
		src_wb.Worksheets.each{ |ws|
		
			# 新規ブックを作成
			dst_wb = excel.workbooks.add()

			# 最初に作成されたワークシート名を変更しておく
			(1..3).each{|num|
				dst_wb.worksheets("Sheet#{num}").name = "#{num}#{num}#{num}#{num}#{num}"
			}

			# ワークシートをコピー
			Excel.sheetCopy( src_wb, ws.name, dst_wb, 1 )
			
			# 最初に作成されたワークシートを削除
			(1..3).each{|num|
				dst_wb.worksheets("#{num}#{num}#{num}#{num}#{num}").delete()
			}

			# シートを保護
			if( is_ws_protect )
				dst_wb.Worksheets(1).Protect
			end

			# 名前をつけて保存
			out_path = out_dir + "/" + "#{ws.name}#{ext_name}"
			
			if( ext_name == ".xls" )
				dst_wb.saveAs( "#{fso.GetAbsolutePathName( out_path )}", -4143 )
			else
				dst_wb.saveAs( "#{fso.GetAbsolutePathName( out_path )}" )
			end
			dst_wb.close(0)				

			# ログ用
			puts "output => " + "#{File.basename(out_path)}"
		}
	end	
end