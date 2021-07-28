"""Detects text in the file."""
import xlsxwriter
import io
from PIL import Image
import os
from os import listdir
from os.path import isfile, join
import sys
from google.cloud import vision

class Chk(object):
	num = -1
	date = []
	to = []
	fore = []
	total = []
		
def make_chk(num, date, to, fore, total):
	chk = Chk()
	chk.num = num
	chk.date = date
	chk.to = to
	chk.fore = fore
	chk.total = total
	return chk

#Seperates the 3 checks into individual checks
def get_checks(ImgSource):
	original = Image.open(ImgSource)
	width, height = original.size
	left = 0
	top = 0
	right = width
	bottom = height
	cropped_temp = original.crop((left+753,top,right,bottom/3))
	check1 = io.BytesIO()
	cropped_temp.save(check1, format = "png")
	check2 = io.BytesIO()
	cropped_temp = original.crop((left+753,top+(bottom/3),right,2*bottom/3))
	cropped_temp.save(check2, format = "png")
	check3 = io.BytesIO()
	cropped_temp = original.crop((left+753,top+(2*bottom/3),right,bottom))
	cropped_temp.save(check3, format = "png")
	return check1, check2, check3
#Get check Number, Date, To, For, Total
def get_chkBoxs(ImgCheck):
	original = Image.open(ImgCheck)
	checkNum = io.BytesIO()
	checkDate = io.BytesIO()
	checkTo = io.BytesIO()
	checkFor = io.BytesIO()
	checkTotal = io.BytesIO()
	width, height = original.size
	left = 0
	top = 0
	right = width
	bottom = height
	cropped_temp = original.crop((254, 57,right-1360, 57+85))
	cropped_temp.show()
	cropped_temp.save(checkNum, format="png")
	cropped_temp = original.crop((154, 125,154+500, 125+106))
	cropped_temp.save(checkDate, format="png")
	cropped_temp = original.crop((50, 215,50+780, 225+205))
	cropped_temp.save(checkTo, format="png")
	cropped_temp = original.crop((50, 425,50+645, 425+295))
	cropped_temp.save(checkFor, format="png")
	cropped_temp = original.crop((875, 510,875+310, 510+105))
	cropped_temp.save(checkTotal, format="png")
	return checkNum, checkDate, checkTo, checkFor, checkTotal
#analyze handwriting add text to array
def analyze_check(checkImg):
	paragraphs = []
	lines= []
	client = vision.ImageAnnotatorClient()
	content = checkImg.getvalue()
	image = vision.types.Image(content=content)

	breaks = vision.enums.TextAnnotation.DetectedBreak.BreakType
	image_context = vision.types.ImageContext(language_hints =["en"])
	response = client.document_text_detection(image=image, image_context = image_context)

	useless_words = {
		"DEPOSITS",
		"TOTAL",
		"OTHER",
		"TAX",
		"DEDUCTIBLE",
		"BALANCE",
		"BAL.",
		"BAL",
		"BRO'T",
		"FOR'D",
		"THIS",
		"CHECK",
		"on",
		"DE POSITS"
	}
	for page in response.full_text_annotation.pages:
			for block in page.blocks:
				
				for paragraph in block.paragraphs:
					para = ""
					line = ""

					for word in paragraph.words:
						for symbol in word.symbols:
							line += symbol.text
							if symbol.property.detected_break.type == breaks.SPACE:
								line += ' '
							if symbol.property.detected_break.type == breaks.EOL_SURE_SPACE:
								line += ' '
								if line.strip() not in useless_words:
									lines.append(line)
									para += line
									line = ''
							if symbol.property.detected_break.type == breaks.LINE_BREAK:
								if line.strip() not in useless_words:
									lines.append(line)
									para += line
									line = ''
				if para != '':
					paragraphs.append(para)

	return lines
#put text from checks into spread sheet
def make_xcel(All_Checks):
	workbook = xlsxwriter.Workbook('Check_Stubs.xlsx')
	worksheet = workbook.add_worksheet('Raw Data Checks')
	row = 1
	tmp = ''
	worksheet.write(0, 0, "Check Number")	
	worksheet.write(0, 1, "Check Date")
	worksheet.write(0, 2, "Check to")	
	worksheet.write(0, 3, "Check for")
	worksheet.write(0, 4, "Check total")	
	for chk in All_Checks:
		worksheet.write(row, 0, chk.num)
		worksheet.write(row, 1, list_str(chk.date))
		worksheet.write(row, 2, list_str(chk.to))	
		worksheet.write(row, 3, list_str(chk.fore))
		worksheet.write(row, 4, list_str(chk.total))	
		row += 1
	workbook.close()
def list_str(list):
	tmp = ''
	for x in list:
		tmp += x
	return x
def main():
	path_to_images = input("Enter path to scanned checks: ")
	all_checks = []
	for path in os.listdir(path_to_images):
		full_path = os.path.join(path_to_images, path)
		if os.path.isfile(full_path):
			all_checks.append(full_path)
	
	All_Checks = []
	for img_chk in all_checks:
		check1 = Chk()
		check2 = Chk()
		check3 = Chk()
		chk1, chk2, chk3 = get_checks(img_chk)
		chkNum, chkDate, chkTo, chkFor, chkTotal = get_chkBoxs(chk1)
		check1.num = int(analyze_check(chkNum)[0].strip())
		check1.date = analyze_check(chkDate)
		check1.to = analyze_check(chkTo)
		check1.fore = analyze_check(chkFor)
		check1.total = analyze_check(chkTotal)
		chkNum, chkDate, chkTo, chkFor, chkTotal = get_chkBoxs(chk2)
		check2.num = int(analyze_check(chkNum)[0].strip())
		check2.date = analyze_check(chkDate)
		check2.to = analyze_check(chkTo)
		check2.fore = analyze_check(chkFor)
		check2.total = analyze_check(chkTotal)
		chkNum, chkDate, chkTo, chkFor, chkTotal = get_chkBoxs(chk3)
		check3.num = int(analyze_check(chkNum)[0].strip())
		check3.date = analyze_check(chkDate)
		check3.to = analyze_check(chkTo)
		check3.fore = analyze_check(chkFor)
		check3.total = analyze_check(chkTotal)
		All_Checks.append(check1)
		All_Checks.append(check2)
		All_Checks.append(check3)
	make_xcel(All_Checks)
main()