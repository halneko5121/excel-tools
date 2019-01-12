# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="

# ==========================="
# src
# ==========================="
class SplitWork
	PARAM_CHECK_WS_NAME	 	= "必要なシート名"
	PARAM_NUMBER					= "番号割振"
	PARAM_CREATE_CALENDAR	= "作成年月"
	PARAM_POST_DIR_NAME			= "postフォルダ名"

	public
	def initialize()
		@check_ws_list			= Array.new
		@check_ws_list.clear
	end

	def getWorkSheetList()
		return @check_ws_list
	end

	def setData( ws_param_works )

		# 列番号の設定
		setClumn( ws_param_works )

		# レコードの数だけ
		for recode in ws_param_works.UsedRange.Rows do 
		
			# 1行目はパラメータ名なのでスキップ or 空白行 or nil が入ってきた場合はスキップ
			next if (recode.row == 1 or recode == "" or recode == nil)

			check_ws_name = Excel.getCellValue(ws_param_works, recode.row, "#{@clumn_check_ws_name}".to_i)
			
			next if (check_ws_name == "" or check_ws_name == nil)
			
			# パラメータを取得してpush
			hash_ws_param = Hash.new
			hash_ws_param[:check_ws_name]	= "#{check_ws_name}"
			temp											= Excel.getCellValue(ws_param_works, recode.row, "#{@clumn_number}".to_i)				
			hash_ws_param[:number]				= temp.to_i
			hash_ws_param[:create_calendar]	= Excel.getCellValue(ws_param_works, recode.row, "#{@clumn_create_calendar}".to_i)
			hash_ws_param[:post_dir_name]		= Excel.getCellValue(ws_param_works, recode.row, "#{@clumn_post_dir_name}".to_i)
			@check_ws_list.push( hash_ws_param )
		end
	end
	
	def split( excel, src_wb, out_dir, ext_name, ws_param_works, is_ws_protect, protect_pass )
	
		# 有給管理表用のパラメータを設定
		setData( ws_param_works )

		# ワークシート名を配列にセットしておく
		ws_name_array = Array.new
		src_wb.Worksheets.each{ |ws|
			ws_name_array.push( ws.name )
		}
		
		# セットされたパラメータの数だけ
		@check_ws_list.each{|data|
	
			# 指定された名前のワークシートがありません
			spilt_ws_name = "#{data[:check_ws_name]}"
			if( ws_name_array.include?( spilt_ws_name ) == false )
				log_str = "シート名に[ #{spilt_ws_name} ]が見当たりません!!"
				assertLogPrintFalse( log_str )
			end		
		}

		fso = WIN32OLE.new('Scripting.FileSystemObject')

		# 該当するワークシートの数だけブックを作成する
		ws_name_array.each{ |ws_name|
		
			# セットされたパラメータの数だけ
			@check_ws_list.each{|data|
			
				if( "#{data[:check_ws_name]}" != ws_name )
					next
				end
				
				# 新規ブックを作成
				dst_wb = excel.workbooks.add()

				# 最初に作成されたワークシート名を変更しておく
				(1..3).each{|num|
					dst_wb.worksheets("Sheet#{num}").name = "#{num}#{num}#{num}#{num}#{num}"
				}
				
				# ワークシートをコピー
				Excel.sheetCopy( src_wb, ws_name, dst_wb, 1 )
				
				# 最初に作成されたワークシートを削除
				(1..3).each{|num|
					dst_wb.worksheets("#{num}#{num}#{num}#{num}#{num}").delete()
				}

				# シートを保護
				if( is_ws_protect && "#{protect_pass}" != "" )
					dst_wb.Worksheets(1).Protect( {'Password' => "#{protect_pass}"} )
				end
				
				dst_path = genarateOutPath( out_dir, data, ext_name )
				save( fso, dst_path, dst_wb, ext_name )
				dst_wb.close(0)				

				# ログ用
				if( is_ws_protect && "#{protect_pass}" != "" )
					print "output => " + "#{File.basename(dst_path)}"
					puts "  protect_pass:#{protect_pass}"
				else
					puts "output => " + "#{File.basename(dst_path)}"
				end
				next				
			}
		}
	end	

	private
	def setClumn( ws )
		@clumn_check_ws_name	= Excel.getColumn(ws, PARAM_CHECK_WS_NAME)
		@clumn_number				= Excel.getColumn(ws, PARAM_NUMBER)
		@clumn_create_calendar	= Excel.getColumn(ws, PARAM_CREATE_CALENDAR)
		@clumn_post_dir_name		= Excel.getColumn(ws, PARAM_POST_DIR_NAME)
	end

	def genarateOutPath( out_dir, data, ext_name )

		staff_num	        = "%03d" % data[:number]
        check_ws_name    = "#{data[:check_ws_name]}".encode( Encoding::UTF_8 ) 
		ws_name         	= "残日数管理#{staff_num}_#{check_ws_name}#{data[:create_calendar]}#{ext_name}".encode(Encoding::Windows_31J)
		out_path        	    = out_dir + "/" + "#{ws_name}"
		return out_path
	end
	
	# 名前をつけて保存
	#ファイル名ルール：残日数管理(0から始まる社員番号)_(シート名)(作成年月日).xlsx
	# *例として2014/1/23　村山　=>　「残日数管理035_村山260123.xlsx」となります。
	def save( fso, save_path, save_wb, ext_name )

		if( ext_name == ".xls" )
			save_wb.saveAs( "#{fso.GetAbsolutePathName( save_path )}", -4143 )
		else
			save_wb.saveAs( "#{fso.GetAbsolutePathName( save_path )}" )
		end	
	end
	
end