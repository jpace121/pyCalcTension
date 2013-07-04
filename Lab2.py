#!/usr/bin/env python

import xlrd as xl
import numpy as np
import matplotlib.pyplot as mpl

class Run(object):
	"""Stores variables for each run for easy access.
	   Allows fake namespace/struct like thing."""
	
	def __init__(self, numRows):
		self.size = numRows
		self.time = [None]*self.size
		self.position = [None]*self.size
		self.velocity = [None]*self.size
		self.fit= [None]*2

	def findVelocity(self):
		"""Given Run object with positions and times, return correpsonding 
			velocity"""
		#Most of the algorithm is completely stolen from the Lab Manual.
		velocity = [None]*self.size
		run = self
		for i in range(run.size+1):
			if i == 0:
				velocity[i] = (run.position[i+1] - run.position[i])/ \
				(run.time[i+1] - run.time[i])
			elif i == (run.size+1):
				velocity[i] = (run.position[i] - run.position[i-1])/ \
				(run.time[i] - run.time[i-1])
			else:
				velocity[i] = (run.position[i+1] - run.position[i-1])/ \
				(run.time[i+1] - run.time[i-1])
			print "i = " + str(i)
			print "Velocity = " + str(velocity[i])
			

	def findAcceleration(self):
		"""Calculate the acceleration and y intercept from the time 
		and velocity variables, add return the linfit"""
		run = self
		time = run.time
		time.pop()
		run.fit = np.polyfit(time, run.velocity, 1)
			

	def plotAcceleration(self):
		"""Plots velocity vs time"""
		run = self
		fit_fun = np.poly1d(run.fit)
		time = run.time
		mpl.plot(time,run.velocity,'bs')
		mpl.show()


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
	excelFile = "ExcelTest.xls"
	sheetName = u"Sheet1"
	Run1 = exportData(excelFile,sheetName)
	Run1.findVelocity()
	Run1.findAcceleration()
	Run1.plotAcceleration()
	print "Acceleration = %d" % Run1.linspace[0]




if __name__ == "__main__":
	main()
