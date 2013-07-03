#!/usr/bin/env python

import xlrd as xl

class Run(object):
	"""Stores variables for each run for easy access.
	   Allows fake namespace/struct like thing."""
	def __init__(self, numRows):
		self.size = numRows
		self.time = [None]*self.size
		self.position = [None]*self.size

def main():
	"""Completes data analysis for Moore's Lab 2, Summer 2013"""
	book = xl.open_workbook('ResultsLab2.xlsx')
	
	run1Sheet = book.sheet_by_name(u'Run1')
	run1 = Run(run1Sheet.nrows)

	#Extract the position values for run1.
	for i in range(run1Sheet.nrows):
		run1.position[i] = run1Sheet.cell(i,0).value
	
	#Extract the time value for run1
	for i in range(run1Sheet.nrows):
		run1.time[i] = run1Sheet.cell(i,1).value

if __name__ == "__main__":
	main()
		
