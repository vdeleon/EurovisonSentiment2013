# -*- coding: utf-8 -*-


import sys
import json
from tempfile import TemporaryFile
from xlsxcessive.xlsx import Workbook
from xlsxcessive.xlsx import save



def sentimentFunct():
	countdown = 0
	sentimentlist = []
	afinnfile = open("AFINN.txt")

	scores = {} # initialize an empty dictionary

	for line in afinnfile:
	  term, score  = line.split("\t")  # The file is tab-delimited. "\t" means "tab character"
	  scores[term] = int(score)  # Convert the score to an integer.


	with open("output.txt") as f:
		x = 0
		
		for line in f:
			while x <= 100:
				if len(line) >= 100:
					sentiment = 0
					wordlist = line.split(" ")
					for word in wordlist:
						if word in scores:
							sentiment += scores[word]
					sentimentlist.append(sentiment)
					x += 1
					print sentiment
				countdown += 1	

				
	return sentimentlist
				


	  
		  
def timestamp():
	countdown = 0
	tweetfile = open("output.txt")
	timestamplist=[]
	
	for tweets in tweetfile:
		if len(tweets) >= 100:
		
			try:
				tweetdict = json.loads(tweets)
		
			except ValueError:
			# decoding failed
				continue
				
			time =  tweetdict["created_at"]	
			zeiteinheiten = time.split(" ")
			zeitunit = zeiteinheiten[3].split(":")
			timestamplist.append(zeitunit)
		
		countdown += 1	
		print "Time %s" % countdown

	return timestamplist

	

def Excel(sentiment, timestamp):	

	book = Workbook(encoding='utf-8')
	sheet1 = book.new_sheet('Sheet 1')
	
	
	sheet1.cell('A1', value='Sentiment')
	sheet1.cell('B1', value="Stunde")
	sheet1.cell('C1', value="Minute")
	sheet1.cell('D1', value="Sekunde")
	
	
	#sentiment schreiben
	countdown = 1
	for einzelbewertung in sentiment:
		countdown +=1
		sheet1.cell("A%s" % countdown, value= einzelbewertung)

		
	
	#zeit hinzufpgen
	countdown = 1
	for zeitangabe in timestamp:
		rower = 0
		celllist= ["B","C","D"]
		countdown += 1
		for einzeleinheit in zeitangabe:
			sheet1.cell("%s%s" % (celllist[rower],countdown), value=einzeleinheit)
			rower += 1
	
	
	book.save('ESC_2117_2300.xlsx')
	book.save(TemporaryFile())
	return "Success"



def main():
	map = {"Sentiment" : sentimentFunct(), "Time": timestamp(), "Excel": Excel(sentimentFunct(),timestamp())}
	return map["Excel"]


if __name__ == '__main__':
    main()