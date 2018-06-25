try:
	from pandas import read_excel	
	from enum import Enum
	from collections import defaultdict
	from collections import namedtuple
	from pprint import pprint
	from sys import exit
	import argparse
	#import json
	import time
	import os
	import xlwt

except ImportError:
	from pandas import read_excel
	from enum import Enum
	from collections import defaultdict
	from collections import namedtuple
	from pprint import pprint
	from sys import exit
	import argparse
	#import json
	import time
	import os
	import xlwt


class State(Enum):
	START = 0
	LOAD = 1
	READ = 2
	WRITE = 3
	ABORT_BY_USER = 4
	ABORT_BY_SYSTEM = 5
	BUILD_INDEX = 6
	SUCCESS = 7
	ERROR = 8	
	EXIT = 9



class MachineState:
	
	states_trans = {
			State.START : { State.ABORT_BY_USER		:	State.EXIT,
							State.ABORT_BY_SYSTEM	:	State.EXIT,
							State.ERROR				:	State.EXIT,
							State.SUCCESS			:	State.LOAD,

						},
			State.LOAD	: { State.ABORT_BY_USER		:	State.EXIT,
							State.ABORT_BY_SYSTEM	:	State.EXIT,
							State.ERROR				:	State.EXIT,
							State.SUCCESS			:	State.READ
						},
			State.READ 	: { State.ABORT_BY_USER		:	State.EXIT,
							State.ABORT_BY_SYSTEM	:	State.EXIT,
							State.ERROR				:	State.EXIT,
							State.SUCCESS			:	State.BUILD_INDEX
						},
			State.BUILD_INDEX
						: { State.ABORT_BY_USER		:	State.EXIT,
							State.ABORT_BY_SYSTEM	:	State.EXIT,
							State.ERROR				:	State.EXIT,
							State.SUCCESS			:	State.WRITE
						},
			State.WRITE	: { State.ABORT_BY_USER		:	State.EXIT,
							State.ABORT_BY_SYSTEM	:	State.EXIT,
							State.ERROR				:	State.EXIT,
							State.SUCCESS			:	State.EXIT
						},
			State.EXIT 	: { }
		}

	def __init__(self):
		self.__state = State.START
		print("Hyy")

	def get_current_state(self):
		return self.__state

	def change_current_state(self,new_state):
		self.__state = self.states_trans[self.__state][new_state]


class DataManager:
		
		def __init__(self, file_name, search_index, position_location, sheet_index = [0,1], build_index = True ):
			"""
			@param file_name -> str or tuple of file names
			@param build_index -> True or False

			"""
			self._state = State.START
			if type(file_name) is type(str):
				self._file_name = tuple(file_name , ) # , due to problem while iterating on tuple
			else: 
				self._file_name = tuple(file_name)
			self._position_location = list(position_location)
			self._file_content = None
			self._sheet_index = sheet_index			
			self._build_index = build_index
			self._index = defaultdict(set)
			self._search_index = search_index
			self._student_data = namedtuple('student',['enrollment', 'index', 'files'])
			

		def set_state(self,new_state):
			self._state = new_state

		#@state(State.READ)
		def read_files(self):

			self.set_state(State.READ)

			assert self._state == State.READ
			if self._file_content is None:
					self._file_content = dict()
			try:
				print(self._file_name)
				for _filename in self._file_name:
					print(_filename)
					for _sheet_index in self._sheet_index:
						print( "Reading {_file}  sheet index = {_sheet}".format(_file = _filename, _sheet = _sheet_index ))
						self._file_content.__setitem__( _filename + str(_sheet_index) , 
															read_excel(_filename,
																			sheet_name = _sheet_index)
														)
				#print(self._file_content)
			except KeyboardInterrupt:
				print("by user")
				raise State.ABORT_BY_USER
			except FileNotFoundError:
				assert self._file_name
				print("file not found {}".format(self._file_name))
				exit()
			#except:
			#	print("by system")
			#	raise State.ABORT_BY_SYSTEM 			
		
		#@state(State.BUILD_INDEX)
		def build_index(self):
			self.set_state(State.BUILD_INDEX)
			
			assert self._state == State.BUILD_INDEX
			
			try:
				# we know all Placement/Training details are on odd index[1,3....] so we will
				# iterate over odd index for building index
				p_i = 0
				for _filename, _content in list(self._file_content.items())[1::len(self._sheet_index)]:
					print("Building index for {_file}".format( _file = _filename ))
					position_index = self._position_location.pop(0)
					
					for index, data in enumerate(_content.values): 
						
						for post in data[ position_index ].split(";"):# data[position_index] is the profiles field
								self._index[post.lower()].add( self._student_data( enrollment = data[1] , 
														index = index,
														 files=p_i) 
														) #data[1] consist of enrolment
					p_i += 1
				#pprint(self._index, indent = 4)
				#with open('index','w+') as f:				
				#	json.dump(self._index, f)
			except KeyboardInterrupt:
				print("by user")
				#raise State.ABORT_BY_USER
				

			#except :
			#	print("by system")
				#raise State.ABORT_BY_SYSTEM


		#@state(State.WRITE)
		def write_files(self):
			#localize the filename for updating
			_file_name = list(self._file_name)
			print("Genrating directory names ")			
			for index, name in enumerate(_file_name):
				name = name.split("/")[-1]
				name = name.replace("Interested in ", "")
				name = name.replace(".xlsx", "")
				_file_name[index] = name
			print("Creating diretory if not exsist")
			for name in _file_name:
				try:
					os.mkdir(name)
				except FileExistsError:
					pass
			
			saved_training = 0 
			saved_placement = 0
			for xlfilename, students in self._index.items():
					placement = xlwt.Workbook() 
					p_sheets = [placement.add_sheet("Acedemic"), 
								placement.add_sheet("Placement")
								]
					training = xlwt.Workbook()
					t_sheets = [training.add_sheet("Acedemic"),
								training.add_sheet("Training")
								]
					row_training = 1 
					row_placement = 1
					for student in students:
						if student.files == 0 : # training 
							for col_index, data in enumerate(
													self._file_content[
														self._file_name[student.files] + "0"].loc(0)[student.index]
															):
								t_sheets[0].write(row_training, col_index, str(data))
							
							for col_index, data in enumerate(
													self._file_content[
														self._file_name[student.files] + "1"].loc(0)[student.index]
															):
								t_sheets[1].write(row_training, col_index, str(data))

							row_training += 1		 
						else: #placement
							for col_index, data in enumerate(
													self._file_content[
														self._file_name[student.files] + "0"].loc(0)[student.index]
															):
								p_sheets[0].write(row_placement, col_index, str(data))
							
							for col_index, data in enumerate(
													self._file_content[
														self._file_name[student.files] + "1"].loc(0)[student.index]
															):
								p_sheets[1].write(row_placement, col_index, str(data))

							row_placement += 1
					
					xlfilename = xlfilename.replace("/", "-")
					xlfilename = xlfilename.replace(" ", "-")
					if row_placement > 1:
						print(f"{xlfilename}.xlsx in {_file_name[1]}{os.sep}")	
						placement.save(f"{_file_name[1]}{os.sep}{xlfilename}.xlsx")
						saved_placement += 1
					if row_training > 1 :
						print(f"{xlfilename}.xlsx in {_file_name[0]}{os.sep}")					
						training.save(f"{_file_name[0]}{os.sep}{xlfilename}.xlsx")
						saved_training += 1
			
			print(f"Total files {saved_placement + saved_training}")
			print(f"Saved {saved_training} files in {_file_name[0]}")
			print(f"Saved {saved_placement} files in {_file_name[1]}") 

		def run(self):
			#try:
				start = time.time()
				self.read_files()
				self.build_index()
				self.write_files()
				finish = time.time() - start
			
				if int(finish) > 3600: 
					print(f"\n Total time taken to complete job {finish/3600} hour {finish / 3600* 60} minutes {finish / 3600*60*60} seconds ")
				elif int(finish) > 60:
					print(f"\n Total time taken to complete job {finish / 60} minutes {finish / 3600} seconds ")
				else:
					print(f"\n Total time taken to complete job {int(finish)} seconds ")
			#except:
			#	print("error")
					

def main():
	parser = argparse.ArgumentParser()
	parser.add_argument("-a","--attribute", help="slect attributes")
	parser.add_argument("-f","--file",type=list,help="file name or directory followed by filename")
	parser.add_argument("-v","--values",type=list, help="values to seprate  eg 1,2,3")
	args = parser.parse_args()

	args.file = [r"/home/anil/Desktop/lmylinux-dir/placement/Interested in Training 2019.xlsx",
				 r"/home/anil/Desktop/lmylinux-dir/placement/Interested in Placement 2019.xlsx" ]
				 
	print(args.file)
	time.sleep(2)

	if args.file is None :
		print("""please include all attributes\n
			 python3 seprator.py -a <attribute_name> -f <file_name> -v <values>
			""")
		return 
	
	"""
	@param 
		@file_name -> consisit of 1 or more files for parsing
		
		@search_index -> are the sheet index of excel files
		
		@position_location -> is for the location of "Position you will like to apply for " file wise
			supply (position you will like to aplly for) index 
			7 is the index of pos.. in file 1			
			31 is the index of position.... in file 2 
			  

	"""	
	datamanager = DataManager(file_name = list(args.file), search_index = [0, 1], position_location=[7, 31] ) 
	
	datamanager.run()
	
	

if __name__ == "__main__":
	main()
	



