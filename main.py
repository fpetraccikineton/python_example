# coding: utf-8
## @author Francesco Petracci
## @version 0.2
## @date 11/10/22
## Goal of this script is to highlight Python possibilities even in the automotive field
## Workflow of the script is as follows:
##      1. load a .dbc file
##      2. extract some info about can messages
##      3. write a table with those messages
##      4. reopen the table and rearrange it to look as I want


import cantools             	# can library, docs here: https://cantools.readthedocs.io/en/latest/
import pandas as pd         	# database analysis lib, with this import I rename the namespace
import os						# module to handle operating specific path/folders
from pathlib import Path		# path lib, we import only the Path function
from support_library import *   # the lib that supports this script, we import everything


script_dir = Path( __file__ ).parent.absolute() # returns the path of the script folder, var ___file___ refers to this script
file_dir = os.listdir(script_dir)	# create a list of files in the directory


list_dbc = [] # init of dbc list
for ele in file_dir:
	if ele[-4:] == ".dbc":
		print("Found a dbc file! " + ele)
		list_dbc.append(ele)

# parse the dbc
msg_name_list 	= []
msg_id_list 	= []
i = 0
for dbc_path in list_dbc:
	print("Parsing " + dbc_path)
	db = cantools.database.load_file(dbc_path)
	
	msg_name_list.append([])
	msg_name_list[i].append(dbc_path)
	msg_id_list.append([])
	msg_id_list[i].append(dbc_path)

	for msg in db.messages:
		msg_name_list[i].append(msg.name)
		msg_id_list[i].append(msg.frame_id)

	i = i + 1

# HOW IN C SHOULD HAVE LOOKED!!
# msg = ""
# for (i, i<n, i++){
# 	msg = db.message(i);
# 	msg ...
# }

# save to excel
excel_path = "dbc_parsed.xlsx"
df = pd.DataFrame( (msg_name_list[0],msg_id_list[0], msg_name_list[1], msg_id_list[1]) )
writer = pd.ExcelWriter(excel_path, engine='openpyxl')
df.to_excel(writer, sheet_name='Test', header=False, index=False)
writer.save()

# wrong orientation! let's change it
invert_row_column(excel_path, "dbc_parsed_right.xlsx")

quit()

#This code will not be executed!

for dbc_path in list_dbc:
	db = cantools.database.load_file(dbc_path)
	
	# print all message ID, and the signals from AIRBAG1 msg
	for msg in db.messages:
		print("%X" %(msg.frame_id))
	print("-----------------------------------------------------")
	try:
		AIRBAG1_msg = db.get_message_by_name('AIRBAG1')
		for sign in AIRBAG1_msg.signals:
			print(sign)
	except:
		print("No AIRBAG1 in " + dbc_path)