#!/usr/bin/env python

import xlrd as xl
import numpy as np
import matplotlib.pyplot as mpl

class Run(object):
	"""Variables and Methods to analyze data from a spreadsheet with times and
	   positions"""
	
	def __init__(self, numRows):
		self.size = numRows
		self.time = [None]*self.size
		self.position = [None]*self.size
		self.velocity = [None]*self.size
		self.fit= []
		self.fitDirect = []
		self.goodVelocities = []
		self.goodTimes = []
		self.goodPositions = []

	def findVelocity(self):
		"""Given Run object with positions and times, return correpsonding 
			velocity"""
		#Most of the algorithm is completely stolen from the Lab Manual.
		for i in range(self.size):
			if i == 0:
				self.velocity[i] = (self.position[i+1] - self.position[i])/ \
				(self.time[i+1] - self.time[i])
			elif i == (self.size-1):
				self.velocity[i] = (self.position[i] - self.position[i-1])/ \
				(self.time[i] - self.time[i-1])
			else:
				self.velocity[i] = (self.position[i+1] - self.position[i-1])/ \
				(self.time[i+1] - self.time[i-1])
	
	def plotVelocity(self):
		"""Plots velocity vs time"""
		fit_fun = np.poly1d(self.fit)
		mpl.plot(self.time,self.velocity,'bs')
		mpl.show()

	def fitAcceleration(self):
		"""Calculates the acceleration directly from the position and time data
		and numpy's linfit option"""
		#Use the calcualted goodTimes to find goodPositions
		for i in range(len(self.goodTimes)):
			self.goodPositions.append(self.position[i])
		self.fitDirect = np.polyfit(self.time, self.position, 2)


def exportData(fileName, sheetName):
	"""Returns Run object, given the name of the excel sheet and the name of
	   the sheet"""
	book = xl.open_workbook(fileName)
	
	runSheet = book.sheet_by_name(sheetName)
	sheetSize = runSheet.nrows - 1
	run = Run(sheetSize)

	#Extract the position values for run1.
	for i in range(run.size):
		run.position[i] = runSheet.cell(i,0).value
	
	#Extract the time value for run1
	for i in range(run.size):
		run.time[i] = runSheet.cell(i,1).value

	return run

def main():
	"""Performs data analysis for Moore's Lab 2."""
	excelFile = "Exceltest.xls"
	sheetName = u"Sheet1"
	Run1 = exportData(excelFile,sheetName)
	Run1.findVelocity()
	#Run1.findAcceleration()
	Run1.fitAcceleration()
	acceleration = Run1.fitDirect[0]*2
	print "Acceleration = %d" % acceleration
	Run1.plotVelocity()



if __name__ == "__main__":
	main()
