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
# def convertKanaFULLWIDTH(transliteration, kanatxt):  # カナ名を半角からスペースを削除して全角にして返す。
# 	transliteration.loadModuleNew((HALFWIDTH_FULLWIDTH,), Locale(Language = "ja", Country = "JP"))
# 	kanatxt = kanatxt.replace(" ", "")  # 半角空白を除去してカナ名を取得。
# 	return transliteration.transliterate(kanatxt, 0, len(kanatxt), [])[0]  # ｶﾅを全角に変換。



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
