import ps
from ps import model
from optparse import OptionParser
from datetime import date, datetime, timedelta
import xml.etree.ElementTree as ET
import sys, os
import logging
import csv
from psPythonUtil import *
import gviz_api
import threading
import webbrowser
import time

def initalizeMe():
	solveModes = ('solve', 'repair')
	usage = "usage: %prog [options]"
	parser = OptionParser(usage=usage)
	parser.add_option("-m", "--mode", dest="solveMode", help="solve or repair")
	options, args = parser.parse_args()
	
	return solveModes, options, args

def getHtml(description, data, psOutputDirectory):
	page_template = """
	<html>
	<head>
	<title>Pack WO Details</title>
	<LINK href="Res/MasterEbs.css" type=text/css rel=stylesheet>
		<script src="http://www.google.com/jsapi" type="text/javascript"></script>
		<script>
		google.load("visualization", "1", {packages:["table"]});
		google.setOnLoadCallback(drawTable);
		function drawTable() {
			var json_table = new google.visualization.Table(document.getElementById('table_div_json'));
			var json_data = new google.visualization.DataTable(%(json)s, 0.5);
			
			var formatter_num = new google.visualization.NumberFormat({'fractionDigits':2});
			var formatter_num00 = new google.visualization.NumberFormat({pattern:'0.000'});
			formatter_num.format(json_data, 4)
			formatter_num.format(json_data, 2)
			formatter_num.format(json_data, 5)
			formatter_num.format(json_data, 7)
			formatter_num.format(json_data, 6)
			formatter_num00.format(json_data, 8)
			formatter_num00.format(json_data, 9)
			formatter_num.format(json_data, 1)
			
			json_table.draw(json_data, {showRowNumber: true, width: 3000, height: 1000});
			
		}
		</script>
	</head>
	<body>
		<div id="header"> 
			<img id="logo" src="Res/FNDSSCORP.gif" alt="Oracle Logo" width="175" height="20" border="0">
			<p id="title">Production Scheduling</p>
		</div>
		<H1 style="text-align:center">Pack Work Order Deltas</H1>
		<div id="table_div_json"></div>
	</body>
	</html>
	"""
	# Creating the data	
	# Loading it into gviz_api.DataTable
	data_table = gviz_api.DataTable(description)
	data_table.LoadData(data)
	
	# Creating a JSon string.   **Need to remove hardcoded columns...todo
	json = data_table.ToJSon(columns_order=("Work Order", "Orig Produced Qty","Adj Produced Qty", "Produced Item", "Operation", "Consumed Item", "Remaining Qty", "Actual Qty", "Delta", "% Delta"),
							order_by="Delta")
	
	# Putting the JS code and JSon string into the template
	#print(page_template % vars())
	htmlFile = os.path.join( psOutputDirectory, 'woDeltas.htm')
	with open(htmlFile, 'w') as f:
		f.write(page_template % vars())
		f.close()


def writeCsv ( list, filename, outDir, order = None ):
	file = filename + '.csv'
	csvFile = os.path.join( outDir, file)
	with open( csvFile, 'wb' ) as f:
		header = []
		for h in list[0].keys():
			header.append( h )
		if order:		#Changes order of the fields based on the list order
			header[:] = [header[i] for i in order]
			
		writer = csv.DictWriter(f, fieldnames=header)
		writer.writeheader()
		for i in list:
			writer.writerow(i)

		f.close()	

def	psGetWoOp(model, threshold, psOutputDirectory):
	workOrders = model.getWorkOrders()				### returns dict {workOrderCode: woObject}
	woDiff=[]
	modifiedWo = model.getAttributes()['ModifiedWO']
	modifiedYes = modifiedWo.findValue('Yes')	
	
	for w, wo in workOrders.iteritems():
		op=wo.getOperations()
		for o in op:
			items = o.getWOConsumedItems()
			opIn = o.getOperationInstances()
			for i in items:
				if (i.actualQuantity-i.quantityRemaining) > threshold and wo.onHold == 0:				
					woDiff.append({ \
					'Work Order': wo.code, \
					'Orig Produced Qty': wo.quantity, \
					'Produced Item': wo.item.code, \
					'Operation':o.code, \
					'Consumed Item': i.item.code, \
					'Remaining Qty':round(i.quantityRemaining,4), \
					'Actual Qty': round(i.actualQuantity,4), \
					'Delta':(round(i.actualQuantity, 4)- round(i.quantityRemaining, 4)), \
					'% Delta': (round(i.actualQuantity, 4)- round(i.quantityRemaining, 4))*100/round(i.actualQuantity,4), \
					'Adj Produced Qty': round((wo.quantity*(1+(round(i.actualQuantity, 4)- round(i.quantityRemaining, 4))/ \
													round(i.actualQuantity,4))), 4)})
					
					for oo in opIn:
						oo.changeAttributeValue(modifiedYes)				

	if woDiff:
		writeCsv (woDiff, 'woDiff', psOutputDirectory, order=[8, 9, 0, 3, 6, 7, 5, 1, 4, 2])  #orders re-arranges the order of the fields i.e WorkOrder is the 5th field but want it to be the first
				
	headers = {
		"Orig Produced Qty": "number",
		"Actual Qty": "number",
		"% Delta": "number",
		"Adj Produced Qty": "number",
		"Delta": "number",
		"Remaining Qty": "number",
		"Operation": "string",
		"Consumed Item": "string",
		"Work Order": "string",
		"Produced Item":"string"
	}
	
	
	
	getHtml(headers, woDiff, psOutputDirectory)

	
	'''
				The following pegs the upstream operation instances (make jobs) so it can be determined which make wo has leftover 
	'''
	'''
			for oo in opIn:
				print "\n OP_INSTANCE:", oo.code, oo.span
				ii = oo.getItemInstances()
				for i in ii:
					print " -ITEMS:", i.item.code, i.quantity
	'''

			
 	return workOrders
	


def open_browser():
	"""Start a browser after waiting for half a second."""
	FILE = 'woDeltas.htm'
	PORT = 8080
	def _open_browser():
		webbrowser.open('http://localhost:%s/%s' % (PORT, FILE))
	thread = threading.Timer(0.5, _open_browser)
	thread.start()


if __name__ == "__main__":	
	configFile = os.path.join( os.getcwd(), 'psPythonConfig.xml')
	variables = setVariables(configFile)
	for key,val in variables.items():
		exec(key + '=val')
	log = setLog(logname, psOutputDirectory)
		
	solveModes, options, args = initalizeMe()
	if options.solveMode is not None and not options.solveMode in solveModes:
		parser.error("Invalid solve mode '" + options.solveMode + " '")
	
	myModel = getModel()	
	if options.solveMode == 'solve' or 'repair':
		psGetWoOp(myModel, float(threshold), psOutputDirectory)

	open_browser()
