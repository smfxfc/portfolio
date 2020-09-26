#! python3
"""portfolio_viz.py - grab market value (cell H7) from portfolio 
history spreadsheet and plot it using matplotlib."""

import matplotlib.pyplot as plt
import openpyxl
from datetime import datetime, timedelta, date

def daterange(start_date, end_date):
	for n in range(int((end_date - start_date).days)):
		yield start_date + timedelta(n)

filename = 'portfolio_history.xlsx' 

wb = openpyxl.load_workbook(filename)

start_date = date(2020, 9, 8)
end_date = datetime.utcnow().date()

dict_vis = {}

for single_day in daterange(start_date, end_date):
	# Catch error for missing dates in portfolio history and skip them
	try:
		active_sheet = wb[single_day.strftime('%y-%m-%d')]
		market_value = active_sheet.cell(7,8).value
		dict_vis[single_day.strftime('%m-%d-%y')] = market_value
	except KeyError:
		pass

x_vals = dict_vis.keys()
y_vals = dict_vis.values()

plt.style.use('classic')
fig, ax = plt.subplots(figsize=(15,9))

#ax.scatter(x_vals, y_vals, s=5)
plt.plot(x_vals, y_vals)
plt.ylim(4000, 6000)

plt.show()