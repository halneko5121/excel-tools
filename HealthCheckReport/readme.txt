==================================
@file			健康チェックシート作成簡略化スクリプト
@auther			wataru mikami
@last_update	2019/01/31
==================================
1.概要
2.使用方法。主な流れ。
3.注意点

1.概要
==================================
　健康チェックリストを、必要人数分用意する手間を簡略化するためのスクリプトです。
　社員が増えた、対象人数が多い。等の要因によって
　総務の方の作業負担が大きく変化する事のないよう
　作業負担軽減が主な目的です。


2.使用方法。主な流れ
==================================
　a.
　--------------------------------------------
　Tools\Template フォルダの中にある
　「Template.xlsx」に問題がないか確認下さい。
　この excel をテンプレートとして、必要分の健康チェックシートが作成されます。

　b.
　--------------------------------------------
　Tools\Template フォルダの中にある
　「TemplateParam.xls」の設定に間違いがないかご確認下さい。
　各社員ごとの情報を、パラメータとして設定する事が可能です。

　c.
　--------------------------------------------
　「a」「b」に問題がなければ 「Template.bat」 を実行して下さい。
　パラメータに設定されている社員分の「健康チェックリスト」が作成されます。


3.注意点
==================================
　テンプレートexcelの、シート名を変更した。列を追加した。などなど
　フォーマットが変更される場合は担当者にご相談下さい。
