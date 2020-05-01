import ps
from ps import model
from optparse import OptionParser
from datetime import date, datetime, timedelta
import xml.etree.ElementTree as ET
import sys, os
import logging
import csv
from psPythonUtil import *

def initalizeMe():
	solveModes = ('solve', 'repair')
	usage = "usage: %prog [options]"
	parser = OptionParser(usage=usage)
	parser.add_option("-m", "--mode", dest="solveMode", help="solve or repair")
	options, args = parser.parse_args()
	
	return solveModes, options, args

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

def	psGetWoOp(model, threshold):
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
					'WorkOrder': wo.code, \
					'WoOriginalProducedQty': wo.quantity, \
					'WOProducedItem': wo.item.code, \
					'Operation':o.code, \
					'ConsumedItem': i.item.code, \
					'RemainingQty':round(i.quantityRemaining,4), \
					'ActualQty': round(i.actualQuantity,4), \
					'Diff':(round(i.actualQuantity, 4)- round(i.quantityRemaining, 4)), \
					'%Percent': (round(i.actualQuantity, 4)- round(i.quantityRemaining, 4))*100/round(i.actualQuantity,4), \
					'WoAdjustedProducedQty': round((wo.quantity*(1+(round(i.actualQuantity, 4)- round(i.quantityRemaining, 4))/ \
													round(i.actualQuantity,4))), 4)})
					
					for oo in opIn:
						oo.changeAttributeValue(modifiedYes)				

	if woDiff:
		writeCsv (woDiff, 'woDiff', psOutputDirectory, order=[8, 9, 0, 3, 6, 7, 5, 1, 4, 2])  #orders re-arranges the order of the fields i.e WorkOrder is the 5th field but want it to be the first
	
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
		psGetWoOp(myModel, float(threshold))
	
