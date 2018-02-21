#docxファイルの文字数、画像数の統計量を計算
#180201追記 ハイパーリンクに未対応。ハイパーリンクは通常のテキストに比べてネストが深くなるため、奥深くまで探索しないといけない。

import xml.etree.ElementTree as ET
import shutil
import zipfile
import os
import sys

#docxファイルの文字数、画像数の統計量を計算
def calsDocxStats(fileheader):
	#docxの中身を見るためにzip形式にし解凍
	shutil.copy(fileheader+".docx", fileheader+".zip")
	with zipfile.ZipFile(fileheader+".zip",'r') as inputFile:
		inputFile.extractall(path=fileheader+"\\")
	#画像ファイル数カウント
	try:
		figure_dirname = fileheader+"/word/media"
		figures_ls = os.listdir(figure_dirname)
		img_n = len(figures_ls)
		#print("画像数は%dです。"%len(figures_ls)) #debug
	except:
		#print("画像はありません。") #debug
		img_n = 0
	#文字数カウント
	xml_filename = fileheader+"/word/document.xml"
	tree = ET.parse(xml_filename)
	root = tree.getroot()
	body = root[0]
	header = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	mojiretsu = ""
	for i in range(len(body)-1):
		line = body[i]
		for ele in line:
			for child in ele:
				#print(child.tag) #debug
				if child.tag == ("{"+header+"}t"): #文字が入っているタグなら文字を描画
					moji = child.text
					mojiretsu += moji
		mojiretsu += "\n"
	#print(moji_retsu) #debug:取得した全文の表示
	moji_n = len(mojiretsu)
	#zipファイルの削除
	os.remove(fileheader+".zip")
	#zip展開ディレクトリの削除
	shutil.rmtree(fileheader)
	return img_n, moji_n

if __name__ == '__main__':
	#計測するファイル名
	docx_filename = input("docxファイル名を入力してください。>")
	fileheader = docx_filename.split(".docx")[0]
	img_n, moji_n = calsDocxStats(fileheader) #docxファイルの文字数、画像数の統計量を計算
	print("画像数は%dです。"%img_n)
	print("文字数は%dです。"%moji_n)