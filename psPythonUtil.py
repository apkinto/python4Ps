import ps
from ps import model
from datetime import date, datetime, timedelta
import xml.etree.ElementTree as ET
import sys, os
import logging
import csv

def getTime():
	currentTime = datetime.now()
	return currentTime

def setVariables(config):
	variable = {}
	config = ET.parse(config)
	root = config.getroot()
	for var in root.find('variableList'):
		variable[var.tag] = var.text
	
	return variable
		
def openModel(inputDir, model, log):
	log.info('\tImport Start...')
	start = getTime()
	psModel = os.path.join(inputDir, model)
	app = ps.app.Application.instance()
	model = app.importModel(psModel)
	end = getTime()
	log.info('\tImport End.  %s sec' % (end-start))
	return model

def	psSolve(model, log):
	log.info('\tSolve Start...')
	start = getTime()
	model.solve()
	end = getTime()
	solveTime = end - start
	log.info('\tSolved in %s sec' % (end - start))

def psSaveDxt(model, outputDir, modelName, log):
	name = modelName.split('.')
	dxtName = name[0] + ".dxt"
	app = ps.app.Application.instance()
	app.saveModel(os.path.join(outputDir, dxtName))

def	psExport(model, outputDir, modelName, log):
	log.info('Exporting...')
	start = getTime()
	name = modelName.split('.')
	xmlName = name[0] + ".xml"
	model.export(os.path.join(outputDir, xmlName))
	end = getTime()
	solveTime = end - start
	log.info('\tExport in %s sec' % (end - start))

'''
#	To be deprecated and replaced with setLog as the current function sets logger >1 resulting in duplicated records written in some cases.

def setLogging():
	logger = logging.getLogger(__name__)
	logger.setLevel(logging.DEBUG)
	

	fh = logging.FileHandler('psPython.log')
	fh.setLevel(logging.INFO)
	
	ch = logging.StreamHandler()
	ch.setLevel(logging.INFO)
	
	#formatter = logging.Formatter('%(asctime)s %(levelname)s %(lineno)d \t%(message)s')
	formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s \t %(message)s \t %(lineno)d')
	fh.setFormatter(formatter)
	ch.setFormatter(formatter)
	
	logger.addHandler(fh)
	logger.addHandler(ch)
	
	return logger
'''

def setLog(logName, outputDir):
	logfile = os.path.join(outputDir, logName)
	logger=logging.getLogger(__name__)
	if not len(logger.handlers):
		logger.setLevel(logging.DEBUG)
		handler=logging.FileHandler(logfile)
		now = datetime.now()
		formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s \t %(message)s \t %(lineno)d')
		handler.setFormatter(formatter)
		logger.addHandler(handler)
	
	return logger

def getModel():
	app = ps.app.Application.instance()
	model = app.currentScenario
	
	return model

def closeLogging():
	logging.shutdown()
