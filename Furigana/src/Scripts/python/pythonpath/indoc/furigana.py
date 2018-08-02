#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)


def createDialog(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	
	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
	selection = doc.getCurrentSelection()
	
def createConfigDialog(xscriptcontext):
	
	
	pass


