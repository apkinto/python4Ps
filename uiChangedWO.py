import ps
from ps import model
from datetime import date, datetime, timedelta
import xml.etree.ElementTree as ET
import sys, os
import logging
import csv
from psPythonUtil import *


def	psGetWoOp(model, threshold):
	workOrders = model.getWorkOrders()				### returns dict {workOrderCode: woObject}
		
	for w, wo in workOrders.iteritems():
		op=wo.getOperations()
		for o in op:
			items = o.getWOConsumedItems()
			opIn = o.getOperationInstances()
			for i in items:
				if (i.actualQuantity-i.quantityRemaining) > threshold and wo.onHold == 0:
					log.info('WO:\'%s\'\tOP:\'%s\'\tCONSUMEDITEM:\'%s\'\tQTYREM: %s\t ACTQTY: %s\t DIFF: %s' % (wo.code, o.code, i.item.code,round(i.quantityRemaining,4),round(i.actualQuantity,4), (round(i.actualQuantity, 4)- round(i.quantityRemaining, 4))) )
			
			'''
			for oo in opIn:
				print "\n OP_INSTANCE:", oo.code, oo.span
				ii = oo.getItemInstances()
				for i in ii:
					print " -ITEMS:", i.item.code, i.quantity
			'''

 	return workOrders

def logIt():
	log.info('psModel: %s', threshold)

	
if __name__ == "__main__":	
	#log = setLogging()
	log = setLog()
	variables = setVariables('psPythonConfig.xml')
	for key,val in variables.items():
		exec(key + '=val')	
	
	myModel = getModel()
	psGetWoOp(myModel, float(threshold))


	
