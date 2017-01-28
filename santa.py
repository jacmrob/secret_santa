import sys, os 
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

import openpyxl
import random

# Class: a gifter 
class Santa():

	def __init__(self, name, email, description, black_list):
		self.name = name 
		self.email = email
		self.description = description
		self.black_list = self.set_up_blacklist(black_list, name)
		self.giftee = None 

	def set_up_blacklist(self, blacklist, name):
		if blacklist:
			return blacklist.split(",") + [name]
		else:
			return [name]


# Class: where the work of secret santas takes place 
class NorthPole():

	def generate_santas(self, file):
		print "Generating secret santas!"
		santas = {}
		wb = openpyxl.load_workbook(file)
		ws = wb.active 
		for row in ws.iter_rows():
			cells = [c.internal_value for c in row]
			santa = Santa(*cells)
			santas[santa.name] = santa 
		print "Done."
		return santas 

	def sort_santas(self, santas):
		print "Assigning secret santas..."
		unassigned = set(santas.keys())
		to_match = [(name, set(data.black_list)) for name, data in santas.iteritems()]
		to_match.sort(key=lambda x: len(x[1]))
		to_match.reverse()

		for person, bade in to_match:
			possible = unassigned - bade 
			if len(possible) < 1:
				raise StandardError("Error! Restart")
			else:
				giftee = list(possible)[random.randint(0, len(possible) - 1)]
				unassigned.remove(giftee)
				santas[person].giftee = giftee 

		print "Done."

	def email_santas(self, santas):
		for name, santa in santas.iteritems():
			print "Emailing %s their secret santa!!" % name 
			giftee_name = santa.giftee 
			giftee = santas[giftee_name]
			self._send_email(santa.email, giftee.name, giftee.description)
		print "Done."

	# todo: pull html out as separate file 
	def _send_email(self, to_addr, giftee, description):
		from_addr = 'secret.krampus2016@gmail.com'
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.starttls()
		server.login(from_addr, 'secretkrampus')

		msg = MIMEMultipart()
		msg['From'] = from_addr
		msg['To'] = to_addr
		msg['Subject'] = 'Your Secret Krampus!'

		body = "Merry Krampusnacht! <br></br> You signed up for Secret Krampus 2016, courtesy of Jackie, Sophie, and Mike. <br></br>" \
			   "The wheel of fate (a random number generator) has been spun, and you'll be giving a gift to... <br></br>" \
			   "<b>{0} </b> <br></br>" \
			   "Your guest of honor has granted you this information to help you out: </br></br>" \
			   "<b>{1}</b> <br></br>" \
			   "Remember to keep your trinket under $20, and to come 'round to 74 Alpine Street at 6pm on December 18th for festivities! <br></br>" \
			   "Happy gifting! <br></br>" \
			   "-- The Krampus Bot (i.e. Jackie's Incredible Coding Skills Inc.)".format(giftee, description)

		msg.attach(MIMEText(body, 'html'))
		server.sendmail(from_addr, to_addr, msg.as_string())
		server.quit()


if __name__ == '__main__':
	f = sys.argv[1]  
	np = NorthPole()
	santas = np.generate_santas(f)
	np.sort_santas(santas)
	np.email_santas(santas)



