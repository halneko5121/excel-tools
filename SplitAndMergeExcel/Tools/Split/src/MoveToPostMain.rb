# -*- coding: utf-8 -*-

# ==========================="
# require
# ==========================="
require File.dirname(__FILE__) + "/../../lib/AppModule.rb"
require File.dirname(__FILE__) + "/../../lib/excel.rb"
require File.dirname(__FILE__) + "/../../lib/util.rb"
require File.dirname(__FILE__) + "/MoveToPost.rb"

# ==========================="
# const
# ==========================="
TITLE		= "MoveToPost"
OUT_ROOT	= File.dirname(__FILE__) + "/../out"

if ( __FILE__ == $0 )

	AppModule.main( TITLE ) {

		puts "注意:[x](閉じるボタン)で終了しないで下さい。"
		puts "excel.exe　のプロセスが残る恐れがあります"
		puts "-----------------------------------------------------------"

		# 共有フォルダに接続
		move_to_post = MoveToPost.new( OUT_ROOT )
		move_to_post.connect( ARGV[0], ARGV[1] )
		
		# 先に移動先一覧を出力する
		move_to_post.printMoveFileList()

		# ファイル移動
		move_to_post.moveFile()
	}

end
