# Excel_order_correct_cancel_price.xlsm
Excel vbaでの注文訂正取消と株価を取得するサンプル 

	ファイル名: e_api_注文訂正取消株価サンプル.xlsm
	言語：Excel VBA
	APIバージョン： V4r3で動作確認

ご注意！！ ================================

 本番環境に接続した場合、実際に注文を出します。

 市場で約定した場合、取り消せません。

 十分に注意してご利用ください。

=========================================

1）動作テストを実行した環境

	os:  Windows 10 Pro 21H1
	excel: 2016 64ビット

２）Excelの設定等

	1.EXCELのメニューからファイル－＞オプション－＞リボンのユーザ設定で開発にチェックを入れる。結果メニューに「開発」が表示されます。
	2.開発－＞Visual Basic を選択、標準モジュールに上記「３．提供ファイル」、2,3 のモジュールを追加します。
	3.VBA のメニュー「ツールー＞参照設定」で以下２つをチェック（参照設定）します。
	  ・Microsoft XML, v6.0
	  ・Microsoft Scripting Runtime
	※EXCEL を新規にインストールした状態で動作確認をしておりますので、EXECEL の各種設定等を変更されている場合は初期状態に戻しご利用下さい。

	また、VBA での JSON 文字列解析に VBA-JSON v2.3.1（https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1）（MIT License）を利用しています。


３）APIの利用には事前に立花証券ｅ支店に口座開設が必要です。


４）利用時にシート「設定」に接続先等を入力してください。

	api接続url
	ログインＩＤ
	ログイン用パスワード（第1）
	注文用パスワード（第２）
	sJsonOfmt（json表示形式）
実際のurl、ご自身のユーザーID、パスワード等に変更してください。
変更しない場合、正常に動作しません。
   

５）各プログラムの実行には、肌色のセルに必要項目を入力してください。青色のセルには返信データを一部出力します。
シート「api_応答」に、APIへの送信テキストと返信データを出力します。
シート「結果」に、返信データを簡単に出力します。


６）実行内容は、以下になります。

	シート	オブジェクト名	プロシージャー名
	--- 	Login	func_login
	--- 	Logout	func_logout
	照会	照会_現物_買付可能額	get_CLMZanKaiKanougaku
	照会	照会_信用_新規建可能額	get_CLMZanShinkiKanoIjiritu
	照会	照会_注文約定一覧	get_CLMOrderList
	照会	照会_注文約定一覧_詳細	get_CLMOrderListDetail
	現物	現物_買い注文	order_gen_buy
	現物	現物_取消	cancel_order
	現物	現物_一括取消	cancel_all_order
	現物	現物_訂正	correct_order
	現物	現物_預り株一覧	get_CLMGenbutuKabuList
	現物	現物_売り注文	order_gen_sell
	信用	信用_新規_買い注文	order_shinki_buy
	信用	信用_新規_売り注文	order_shinki_sell
	信用	信用_建玉一覧	get_CLMShinyouTategyokuList
	信用	信用_返済_買い注文	order_hensai_buy
	信用	信用_返済_売り注文	order_hensai_sell
	信用	信用_返済_買い注文_個別指定	order_hensai_buy_aCLMKabuHensaiData
	信用	信用_返済_売り注文_個別指定	order_hensai_sell_aCLMKabuHensaiData
	株価	日足データ取得		get_CLMMfdsGetMarketPriceHistory
	株価	スナップショット	get_CLMMfdsGetMarketPrice
	マスター	マスター_個別　株式銘柄	get_master_kobetu_CLMIssueMstKabu
	マスター	マスター_個別　銘柄市場	get_master_kobetu_CLMIssueSizyouMstKabu
	マスター	マスター_個別　先物	get_master_kobetu_CLMIssueMstSak
	マスター	マスター_個別　OP	get_master_kobetu_CLMIssueMstOp
	マスター	マスター_個別　指数、為替、その他	get_master_kobetu_CLMIssueMstOther
	マスター	マスター_個別　取引所エラー理由コード	get_master_kobetu_CLMOrderErrReason
	マスター	マスター_個別　日付情報	get_master_kobetu_CLMDateZyouhou
 	（現在、マスターデータの更新に不具合が発生しており、誤った情報が記載されています。ご注意ください。）



７）利用時間外に接続した場合、"p_errno":"9"（システム、サービス停止中。）が返されます。詳しくは「立花証券・ｅ支店・ＡＰＩ、EVENT I/F 利用方法、データ仕様」4ページをご参照ください。デモ環境の利用時間は、デモ環境の説明ページを参照ください。

８）本サンプルプログラムは、事務方の者が休日や空き時間に作成したため、色々足りておりません。ご容赦ください。

９）本プログラムはライセンスの範囲内で自由にご利用ください。

１０）このソフトウェアを使用したことによって生じたすべての障害・損害・不具合等に関して、私と私の関係者および私の所属するいかなる団体・組織とも、一切の責任を負いません。各自の責任においてご使用ください。
