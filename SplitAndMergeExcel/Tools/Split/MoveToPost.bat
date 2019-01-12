@set RUBY_LIB_DIR=..\lib\ruby-dist
@set SRC_DIR=.\src
@rem 共有フォルダに接続する場合
@set password=xxxxxxxxxx
@set user_name=xxxxxxxxxx

@%RUBY_LIB_DIR%\ruby.exe %SRC_DIR%\MoveToPostMain.rb %password% %user_name%
pause