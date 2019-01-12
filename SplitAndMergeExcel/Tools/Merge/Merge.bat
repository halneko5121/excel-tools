@set RUBY_LIB_DIR=..\lib\ruby-dist
@set SRC_DIR=.\src

%RUBY_LIB_DIR%\ruby.exe %SRC_DIR%\SrcMain.rb %*
pause