import pandas as pd
import os
import datetime

def save_pivot_table(locations, date, weightlist, volumelist, output):
	pivot=[]
	year = date.date().year
	month = date.date().month

	l=[' ']

	for i in range(1,32):
		try:
			d=datetime.datetime(int(year),int(month),int(i))
			date = str(month)+'/'+str(i)+'/'+str(year)
			l.append(date)
			l.append(date)
		except:
			a=1
	pivot.append(l)

	l=['Location']
	for i in range(1,32):
		try:
			d=datetime.datetime(int(year),int(month),int(i))
			l.append('weight(kg)')
			l.append('volume(m3)')
		except:
			a=1
	pivot.append(l)

	for L in locations:
		l=[L+' Branch']
		for i in range(1,32):
			try:
				d=datetime.datetime(int(year),int(month),int(i))
				if L in weightlist[i]:
					if weightlist[i][L] > 500:
						l.append(weightlist[i][L])
						l.append(volumelist[i][L])
					else:
						l.append(0)
						l.append(0)
				else:
					l.append(0)
					l.append(0)
			except:
				a=1
		pivot.append(l)

	df = pd.DataFrame(pivot)
	df.to_excel(output,index=False)
	print("    Pivot table created. ")


def make_pivot_table(sheet, month):

	if len(sheet.values.tolist())==0:
		return 0

	print("\nStep 2:  Creating pivot tables")

	if 'Output' not in os.listdir():
		os.mkdir('Output')
		
	if month not in os.listdir('Output/'):
		os.mkdir('Output/'+month)

	output_pivot_table = 'Output/'+month+'/PivotTable_'+month+'.xlsx'

	branches = set()
	try:
		BRS_list = pd.read_excel('Master.xlsx', 'BRS').values

		for i in BRS_list:
			br = i[1].split(' + ')
			for b in br:
				branches.add(b.split(' ')[0])
	except:
		a=1

	date_df = sheet['LocTfrDate']
	location = sheet['TOLOCATION']
	weight = sheet['Weight_kg']
	volume = sheet['Volume_m3']
	locations = set()

	date=[]
	for i in date_df:
		if type(i) == type('abc'):
			date.append(datetime.datetime.strptime(i, '%m/%d/%Y'))
		else:
			date.append(i)

	n = len(location)

	weightlist = [dict() for x in range(32)]
	volumelist = [dict() for x in range(32)]


	for i in range(n):
		try:
			loc = location[i].split(' ')[0]
			day = date[i].date().day
			if loc in branches:
				locations.add(loc)
				if loc in weightlist[day]:
					weightlist[day][loc] = weightlist[day][loc] + weight[i]
					volumelist[day][loc] = volumelist[day][loc] + volume[i]
				else:
					weightlist[day][loc] = weight[i]
					volumelist[day][loc] = volume[i]
		except :
			a=1

	branch_sheet = pd.read_excel('Master.xlsx', 'milk run')
	branches = branch_sheet.values

	for b in branches:

		name = ''

		for i in b[1:]:
			try:
				name = name+i.split(' ')[0]+' + '
				locations.remove(i.split(' ')[0])
			except :
				a=1
		name = name[:-3]
		locations.add(name)

		for day in range(1,32):
			weight = 0
			volume = 0
			
			for i in b[1:]:
				try:
					br = i.split(' ')[0]
					if br in weightlist[day]:
						weight=weight+weightlist[day][br]
						volume=volume+volumelist[day][br]

						del weightlist[day][br]
						del volumelist[day][br]
				except :
					a=1

			weightlist[day][name] = weight
			volumelist[day][name] = volume

	save_pivot_table(locations, date[0], weightlist, volumelist, output_pivot_table)


def save_parcel_cost_table(locations, date, caselist, output):
	pivot=[]
	year = date.date().year
	month = date.date().month

	l=[' ', ' ']

	for i in range(1,32):
		try:
			d=datetime.datetime(int(year),int(month),int(i))
			date = str(month)+'/'+str(i)+'/'+str(year)
			l.append(date)
		except:
			a=1
	pivot.append(l)

	l=['Location','Total Cost']
	for i in range(1,32):
		try:
			d=datetime.datetime(int(year),int(month),int(i))
			l.append('Cost')
		except:
			a=1
	pivot.append(l)

	rates = dict()
	Prates = pd.read_excel('Master.xlsx','Parcel Rates').values
	for r in Prates:
		rates[r[1].split(' ')[0].upper()] = r[2]

	for L in locations:
		l=[L+' Branch',0]
		s=0
		for i in range(1,32):
			try:
				d=datetime.datetime(int(year),int(month),int(i))
				if L in caselist[i]:
					l.append(caselist[i][L] * rates[L])
					s+= caselist[i][L] * rates[L]
				else:
					l.append(0)
			except:
				a=1
		l[1]=s
		pivot.append(l)

	df = pd.DataFrame(pivot)
	df.to_excel(output,index=False)
	print("    Parcel cost table created.")


def make_parcel_cost_table(sheet, month):

	if len(sheet.values.tolist())==0:
		return 0

	if 'Output' not in os.listdir():
		os.mkdir('Output')
		
	if month not in os.listdir('Output/'):
		os.mkdir('Output/'+month)

	output_parcel_cost = 'Output/'+month+'/ParcelCost_'+month+'.xlsx' 

	branches = set()
	try:
		parcel_list = pd.read_excel('Master.xlsx', 'Parcel Rates').values

		for i in parcel_list:
			branches.add(i[1].split(' ')[0].upper())
	except:
		a=1

	date_df = sheet['LocTfrDate']
	location = sheet['TOLOCATION']
	case = sheet['Case']
	locations = set()

	date=[]
	for i in date_df:
		if type(i) == type('abc'):
			date.append(datetime.datetime.strptime(i, '%m/%d/%Y'))
		else:
			date.append(i)

	n = len(location)

	caselist = [dict() for x in range(32)]

	for i in range(n):
		try:
			loc = location[i].split(' ')[0]
			day = date[i].date().day
			if loc in branches:
				locations.add(loc)
				if loc in caselist[day]:
					caselist[day][loc] = caselist[day][loc] + case[i]
				else:
					caselist[day][loc] = case[i]
		except :
			a=1

	locations = sorted(locations)
	
	save_parcel_cost_table(locations, date[0], caselist, output_parcel_cost)



