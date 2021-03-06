#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# import os, unohelper
import os
from indoc import sheet, documentevent  # 相対インポートは不可。
# from com.sun.star.awt import MessageBoxButtons  # 定数
# from com.sun.star.awt.MessageBoxType import ERRORBOX  # enum
# from com.sun.star.datatransfer import XTransferable
# from com.sun.star.datatransfer import DataFlavor  # Struct
# from com.sun.star.datatransfer import UnsupportedFlavorException
# from com.sun.star.i18n.TransliterationModulesNew import HALFWIDTH_FULLWIDTH  # enum
# from com.sun.star.lang import Locale  # Struct
# from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
# from com.sun.star.sheet.CellInsertMode import ROWS as insert_rows  # enum
# from com.sun.star.table import BorderLine2, TableBorder2 # Struct
# from com.sun.star.table import BorderLineStyle  # 定数
# from com.sun.star.table.CellHoriJustify import LEFT  # enum
def getModule(sheetname):  # シート名に応じてモジュールを振り分ける関数。
	if sheetname is None:  # シート名でNoneが返ってきた時はドキュメントイベントとする。
		return documentevent
	else:
		return sheet
	

# class TextTransferable(unohelper.Base, XTransferable):
# 	def __init__(self, txt):  # クリップボードに渡す文字列を受け取る。
# 		self.txt = txt
# 		self.unicode_content_type = "text/plain;charset=utf-16"
# 	def getTransferData(self, flavor):
# 		if flavor.MimeType.lower()!=self.unicode_content_type:
# 			raise UnsupportedFlavorException()
# 		return self.txt
# 	def getTransferDataFlavors(self):
# 		return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),  # DataTypeの設定方法は不明。
# 	def isDataFlavorSupported(self, flavor):
# 		return flavor.MimeType.lower()==self.unicode_content_type
# def formatkeyCreator(doc):  # ドキュメントを引数にする。
# 	def createFormatKey(formatstring):  # formatstringの書式はLocalによって異なる。 
# 		numberformats = doc.getNumberFormats()  # ドキュメントのフォーマット一覧を取得。デフォルトのフォーマット一覧はCalcの書式→セル→数値でみれる。
# 		locale = Locale(Language="ja", Country="JP")  # フォーマット一覧をくくる言語と国を設定。インストールしていないUIの言語でもよい。。 
# 		formatkey = numberformats.queryKey(formatstring, locale, True)  # formatstringが既存のフォーマット一覧にあるか調べて取得。第3引数のブーリアンは意味はないはず。 
# 		if formatkey == -1:  # デフォルトのフォーマットにformatstringがないとき。
# 			formatkey = numberformats.addNew(formatstring, locale)  # フォーマット一覧に追加する。保存はドキュメントごと。 
# 		return formatkey
# 	return createFormatKey
# def createBorders():# 枠線の作成。
# 	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
# 	firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=COLORS["violet"])
# 	secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=COLORS["magenta3"])	
# 	tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
# 	topbottomtableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=False, IsRightLineValid=False)
# 	leftrighttableborder = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=False, IsBottomLineValid=False, IsLeftLineValid=True, IsRightLineValid=True)
# 	return noneline, tableborder2, topbottomtableborder, leftrighttableborder  # 作成した枠線をまとめたタプル。
# def convertKanaFULLWIDTH(transliteration, kanatxt):  # カナ名を半角からスペースを削除して全角にして返す。
# 	transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))
# 	kanatxt = kanatxt.replace(" ", "")  # 半角空白を除去してカナ名を取得。
# 	return transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # ｶﾅを全角に変換。
# def createKeikaPathname(doc, transliteration, idtxt, kanatxt, filename):
# 	kanatxt = convertKanaFULLWIDTH(transliteration, kanatxt)  # カナ名を半角からスペースを削除して全角にする。
# 	dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
# 	return os.path.join(dirpath, "*", filename.format(kanatxt, idtxt))  # ワイルドカード入のシートファイル名を取得。	
# def showErrorMessageBox(controller, msg):
# 	componentwindow = controller.ComponentWindow
# 	msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, ERRORBOX, MessageBoxButtons.BUTTONS_OK, "myRs", msg)
# 	msgbox.execute()
# def getKaruteSheet(doc, idtxt, kanjitxt, kanatxt, datevalue):
# 	sheets = doc.getSheets()  # シートコレクションを取得。
# 	if idtxt in sheets:  # すでに経過シートがある時。
# 		karutesheet = sheets[idtxt]  # カルテシートを取得。  
# 	else:
# 		sheets.copyByName("00000000", idtxt, len(sheets))  # テンプレートシートをコピーしてID名のシートにして最後に挿入。	
# 		karutesheet = sheets[idtxt]  # カルテシートを取得。  
# 		karutevars = karute.VARS
# 		karutevars.setSheet(karutesheet)	
# 		karutedatecell = karutesheet[karutevars.splittedrow, karutevars.datecolumn]
# 		karutedatecell.setValue(datevalue)  # カルテシートに入院日を入力。
# 		createFormatKey = formatkeyCreator(doc)
# 		karutedatecell.setPropertyValues(("NumberFormat", "HoriJustify"), (createFormatKey('YYYY/MM/DD'), LEFT))  # カルテシートの入院日の書式設定。左寄せにする。
# 		karutesheet[:karutevars.splittedrow, karutevars.articlecolumn].setDataArray((("",), (" ".join((idtxt, kanjitxt, kanatxt)),)))  # カルテシートのコピー日時をクリア。ID名前を入力。
# 	return karutesheet	
# def getKeikaSheet(doc, idtxt, kanjitxt, kanatxt, datevalue):
# 	sheets = doc.getSheets()  # シートコレクションを取得。
# 	newsheetname = "".join([idtxt, "経"])  # 経過シート名を取得。
# 	if newsheetname in sheets:  # すでに経過シートがある時。
# 		keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
# 	else:	
# 		sheets.copyByName("00000000経", newsheetname, len(sheets))  # テンプレートシートをコピーしてID経名のシートにして最後に挿入。	
# 		keikasheet = sheets[newsheetname]  # 新規経過シートを取得。
# 		keikavars = keika.VARS
# 		keikasheet[keikavars.daterow, keikavars.yakucolumn].setString(" ".join((idtxt, kanjitxt, kanatxt)))  # ID漢字名ｶﾅ名を入力。					
# 		keika.setDates(doc, keikasheet, keikasheet[keikavars.daterow, keikavars.splittedcolumn], datevalue)  # 経過シートの日付を設定。
# 	return keikasheet	
# def toNewEntry(sheet, rangeaddress, edgerow, dest_row):  # 使用中最下行へ。新規行挿入は不要。
# 	startrow, endrowbelow = rangeaddress.StartRow, rangeaddress.EndRow+1  # 選択範囲の開始行と終了行の取得。
# 	if endrowbelow>edgerow:
# 		endrowbelow = edgerow
# 	sourcerangeaddress = sheet[startrow:endrowbelow, :].getRangeAddress()  # コピー元セル範囲アドレスを取得。
# 	sheet.moveRange(sheet[dest_row, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。	
# 	sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動したソース行を削除。
# def toOtherEntry(sheet, rangeaddress, edgerow, dest_row):  # 新規行挿入が必要な移動。
# 	startrow, endrowbelow = rangeaddress.StartRow, rangeaddress.EndRow+1  # 選択範囲の開始行と終了行の取得。
# 	if endrowbelow>edgerow:
# 		endrowbelow = edgerow
# 	sourcerange = sheet[startrow:endrowbelow, :]  # 行挿入前にソースのセル範囲を取得しておく。
# 	dest_rangeaddress = sheet[dest_row:dest_row+(endrowbelow-startrow), :].getRangeAddress()  # 挿入前にセル範囲アドレスを取得しておく。
# 	sheet.insertCells(dest_rangeaddress, insert_rows)  # 空行を挿入。	
# 	sheet.queryIntersection(dest_rangeaddress).clearContents(511)  # 挿入した行の内容をすべてを削除。挿入セルは挿入した行の上のプロパティを引き継いでいるのでリセットしないといけない。
# 	sourcerangeaddress = sourcerange.getRangeAddress()  # コピー元セル範囲アドレスを取得。行挿入後にアドレスを取得しないといけない。
# 	sheet.moveRange(sheet[dest_row, 0].getCellAddress(), sourcerangeaddress)  # 行の内容を移動。			
# 	sheet.removeRange(sourcerangeaddress, delete_rows)  # 移動したソース行を削除。		
# # 	
# # 	
# # 	
# # 	
# 以下コンテクストメニュー
def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	i = 0  # インデックスを初期化する。
	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
		nonlocal i
		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
		i += 1  # インデックスを増やす。
	return addMenuentry
# def cutcopypasteMenuEntries(addMenuentry):  # コンテクストメニュー追加。
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# def rowMenuEntries(addMenuentry):  # コンテクストメニュー追加。
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsAfter"})
# 	addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
def getBaseURL(xscriptcontext):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	filepath = modulepath.replace(ucp, "")  #  ucpを除去。ドキュメントを一旦閉じて開き直してもContentIdentifierが更新されない。
	filepath = os.path.join(*filepath.split("/")[2:])  # Scripts/python/pythonpath/indoc/commons.py。ContentIdentifierを除く。
	macrofolder = "Scripts/python"
	location = "document"  # マクロの場所。	
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	
def invokeMenuEntry(entrynum):  # コンテクストメニュー項目から呼び出された処理をシートごとに振り分ける。コンテクストメニューから呼び出しているこの関数ではXSCRIPTCONTEXTが使える。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	selection = doc.getCurrentSelection()  # セル(セル範囲)またはセル範囲、セル範囲コレクションが入るはず。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲コレクション以外の時。
		m = getModule(doc.getCurrentController().getActiveSheet().getName())
		if hasattr(m, "contextMenuEntries"):
			getattr(m, "contextMenuEntries")(entrynum, XSCRIPTCONTEXT)	
# ContextMenuInterceptorのnotifyContextMenuExecute()メソッドで設定したメニュー項目から呼び出される関数。関数名変更不可。動的生成も不可。
def entry1():
	invokeMenuEntry(1)
# def entry2():
# 	invokeMenuEntry(2)	
# def entry3():
# 	invokeMenuEntry(3)	
# def entry4():
# 	invokeMenuEntry(4)
# def entry5():
# 	invokeMenuEntry(5)
# def entry6():
# 	invokeMenuEntry(6)
# def entry7():
# 	invokeMenuEntry(7)
# def entry8():
# 	invokeMenuEntry(8)
# def entry9():
# 	invokeMenuEntry(9)	
# def entry10():
# 	invokeMenuEntry(10)	
# def entry11():
# 	invokeMenuEntry(11)	
# def entry12():
# 	invokeMenuEntry(12)	
# def entry13():
# 	invokeMenuEntry(13)	
# def entry14():
# 	invokeMenuEntry(14)	
# def entry15():
# 	invokeMenuEntry(15)	
# def entry16():
# 	invokeMenuEntry(16)	
# def entry17():
# 	invokeMenuEntry(17)	
# def entry18():
# 	invokeMenuEntry(18)	
# def entry19():
# 	invokeMenuEntry(19)	
# def entry20():
# 	invokeMenuEntry(20)	
# # ファイルの選択はこれ以降を使用する。ファイルが増えたら単に連番の関数を追加するだけで良い。
# def entry21():
# 	invokeMenuEntry(21)	
# def entry22():
# 	invokeMenuEntry(22)	
# def entry23():
# 	invokeMenuEntry(23)	
# def entry24():
# 	invokeMenuEntry(24)	
# def entry25():
# 	invokeMenuEntry(25)	
# def entry26():
# 	invokeMenuEntry(26)	
# def entry27():
# 	invokeMenuEntry(27)	
# def entry28():
# 	invokeMenuEntry(28)	
# def entry29():
# 	invokeMenuEntry(29)	
# def entry30():
# 	invokeMenuEntry(30)	
# 	