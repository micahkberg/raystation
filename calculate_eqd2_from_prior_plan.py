import datetime
import statetree
from math import ceil
from connect import *
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Application, Form, Label, ComboBox, Button, TextBox
from System.Drawing import Point, Color, Size

class CheckRestTime(Form):
	def __init__(self, rest, treatdate):
		self.Size = Size(450, 150)
		self.rest = rest
		self.treatdate = treatdate
		self.Text = "EQD2 From Plan"
		self.inputbox = None
		label = Label()
		label.Text = f"Is an original treatment date of {treatdate} or a rest time of {rest} acceptable? (y/n)"
		label.Location = Point(15, 15)
		label.AutoSize = True
		self.Controls.Add(label)

		yes_button = Button()
		yes_button.Text = "Yes"
		yes_button.AutoSize = True
		yes_button.Location = Point(15,50)
		yes_button.Click += self.yes_button_clicked

		no_button = Button()
		no_button.Text = "No (manual submit rest time)"
		no_button.AutoSize = True
		no_button.Location = Point(100,50)
		no_button.Click += self.no_button_clicked

		self.Controls.Add(yes_button)
		self.Controls.Add(no_button)

	def yes_button_clicked(self, sender, event):
		self.Close()

	def no_button_clicked(self, sender, event):
		self.Controls.Clear()

		label = Label()
		label.Text = "Submit an approximate original treat date in mm/dd/yyyy format"
		label.Location = Point(15,15)
		label.AutoSize = True
		self.Controls.Add(label)

		self.inputbox = TextBox()
		self.inputbox.Location = Point(15,75)
		self.inputbox.Size = Size(15,100)
		self.inputbox.AutoSize = True
		self.inputbox.Text = "mm/dd/yyyy"
		self.inputbox.ForeColor = Color.Gray
		def highlight_text(sender, event):
			if self.inputbox.Text == "mm/dd/yyyy":
				self.inputbox.Text = ""
				self.inputbox.ForeColor = Color.Black
		def unhighlight_text(sender, event):
			if self.inputbox.Text == "":
				self.inputbox.Test = "mm/dd/yyyy"
				self.inputbox.ForeColor = Color.Gray

		self.inputbox.GotFocus += highlight_text
		self.inputbox.LostFocus += unhighlight_text
		self.Controls.Add(self.inputbox)

		accept_button = Button()
		accept_button.Text = "OK"
		accept_button.Position = Point(150,150)
		accept_button.AutoSize = True
		accept_button.Click += self.accept_button_clicked
		self.Controls.Add(accept_button)

	def accept_button_clicked(self, sender, event):
		response_split = [int(i) for i in self.inputbox.Text.split("/")]
		user_treat_date = datetime.date(response_split[2], response_split[0], response_split[1])
		self.rest = (datetime.date.today() - user_treat_date).days
		self.Close()


plan=get_current('Plan')

def prompt_confirm_rest_time(r, t):
	new_window = CheckRestTime(r, t)
	Application.Run(new_window)
	return new_window.rest



n = plan.BeamSets[0].FractionationPattern.NumberOfFractions #number of fractions
TxTime = n + (ceil(n/5)-1)*2 # in days, calculating number of weekends assuming start treatment date of monday
TxDateObj = plan.Review.ReviewTime
LastTreatTime = datetime.date(TxDateObj.Year, TxDateObj.Month, TxDateObj.Day) #taken from date of plan review
RestTime = prompt_confirm_rest_time((datetime.date.today() - LastTreatTime).days, LastTreatTime)


def get_organ_alphabeta(organ):
	return 2
	#some organs
	#	return 2
	#other organs
	#	return 3

def get_eqd2(roi):
	D = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName=roi, DoseType="Max")/100 #returns d*n, max dose over course of treatment
	d = D/n
	alphabeta = get_organ_alphabeta(roi)
	eqd2 = D * ((alphabeta + d)/(2 +alphabeta)) * (TxTime/(TxTime+RestTime))**0.11
	# equivalent dose in 2 Gy equivalent form
	# D = total max pt dose to OAR in Gy
	# d = fractional max pt dose to OAR in Gy (ie max pt Dose/n)
	return eqd2

def main():
	ROIs = plan.GetTotalDoseStructureSet().RoiGeometries
	table = []
	for roi in ROIs:
		roi_name = str(roi).split(",")[1][2:-2]
		print(roi_name)
		table.append([roi_name,get_eqd2(roi_name)])
	for line in table:
		print(line)

main()
