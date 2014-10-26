# NED to FVS Converter
# Author: Casey Ghilardi
# Website: caseyghilardi.blogspot.com
# Email: Crghilardi@yahoo.com
# NJ State Forestry Services
# Central Region, New Lisbon Office
# December 2013


from Tkinter import Tk
from tkFileDialog import askopenfilename

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

import os
import shutil
DB = 'Blank_DB_copy.mdb'
ned_db = filename 
fvs = '_FVS'
mdb = '.mdb'
 
ned_strip = len(ned_db) - 4 # take away '.mdb'
ned_new = ned_db[0:ned_strip]
new_fname = ned_new + fvs + mdb
new_fname2 = ned_new + fvs


try:
   shutil.copy(DB,new_fname)
except Exception,e:
   print e

cnxn_str = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
cnxn_in = cnxn_str + ned_new
cnxn_out = cnxn_str + new_fname2 

ned_short = os.path.basename(ned_new)
fvs_short = os.path.basename(new_fname2)

stand_header = '.Stand_header'
inv_param = '.Inventory_parameters'
tree_list = '.FVS_TreeInit'
stand_list = '.FVS_StandInit'
ov_obs = '.Overstory_obs'

fvs_stand = fvs_short + stand_list
fvs_tree = fvs_short + tree_list
ned_obs = ned_short + ov_obs

fvs_headers = fvs_short + stand_header
ned_headers = ned_short + stand_header

fvs_param = fvs_short + inv_param
ned_param = ned_short + inv_param


import pyodbc
import time 

##########################################################
# Connections to DBs
### need to add file browser and then tie in to connection string ###

start = time.time()

cnxn = pyodbc.connect(cnxn_in)
cursor = cnxn.cursor()

cnxn2 = pyodbc.connect(cnxn_out)
cursor2 = cnxn2.cursor()


##########################################################
# Add stand header  and inventory parameter tables (contains stand names) to FVS DB **essentially how to copy and paste a table**

Query1 = "SELECT * INTO " + fvs_headers + " FROM " + ned_headers
cursor2.execute(Query1)
cnxn2.commit()

Query2 = "SELECT * INTO " + fvs_param + " FROM " + ned_param
cursor2.execute(Query2)
cnxn2.commit()
##########################################################
#Query out info from NED DB and insert into FVS DB

Query3 = "INSERT INTO " + fvs_tree + "(Stand_ID, Plot_ID, TREE_ID, Tree_Count, Species, DBH, History, DG, HT, Age) " + "SELECT SNAPSHOT, [CLUSTER] + 1, [OVER_OBS] + 1, tree_stems, tree_spp, tree_dbh, tree_alive, tree_user2, tree_height, tree_user1 " + "FROM " + ned_obs
print Query3 
cursor2.execute(Query3)
cnxn2.commit()

##########################################################
#Change data type in stand_headers table to be able to relate them to other tables

Query4 = '''ALTER TABLE Stand_header ALTER COLUMN STAND TEXT'''
cursor2.execute(Query4)
cnxn2.commit()

##########################################################
#spp conversion

Query5 = '''UPDATE FVS_TreeInit INNER JOIN spp_lookup_table ON FVS_TreeInit.Species = spp_lookup_table.PLANTS_Spp
            SET FVS_TreeInit.Species = [spp_lookup_table].[FVS_AlphaCode];'''
cursor2.execute(Query5)
cnxn2.commit()

##########################################################
#Live/Dead conversion

Query6 = '''UPDATE FVS_TreeInit INNER JOIN Live_dead ON FVS_TreeInit.History = Live_dead.NED_LD
            SET FVS_TreeInit.History = [Live_dead].[FVS_History];'''
cursor2.execute(Query6)
cnxn2.commit()


###########################################################
#Populate NED-extractable StandInit table variables

Query7 = "INSERT INTO " +  fvs_stand + "(Stand_ID, Inv_Year, Basal_Area_Factor, Inv_Plot_Size) " + "SELECT SNAPSHOT, '20' + RIGHT(inventory_date,2) , inventory_over_baf, 1 / [inventory_under_size] FROM Inventory_parameters"

cursor2.execute(Query7)
cnxn2.commit()


Query8 = '''UPDATE FVS_StandInit INNER JOIN Stand_header ON FVS_StandInit.Stand_ID = Stand_header.STAND
            SET FVS_StandInit.Sam_Wt = [Stand_header].[stand_area];'''

cursor2.execute(Query8)
cnxn.commit()

################################################################
#Convert SNAPSHOT numbers into stand names based on stand header table
# moved to later so can use snapshot numbers to join tables
Query9 = '''UPDATE FVS_TreeInit INNER JOIN Stand_header ON FVS_TreeInit.Stand_ID = Stand_header.STAND
            SET FVS_TreeInit.Stand_ID = [Stand_header].[stand_id];'''
cursor2.execute(Query9)
cnxn2.commit()

Query10 = '''UPDATE FVS_StandInit INNER JOIN Stand_header ON FVS_StandInit.Stand_ID = Stand_header.STAND
            SET FVS_StandInit.Stand_ID = [Stand_header].[stand_id];'''
cursor2.execute(Query10)
cnxn2.commit()



##########################################################
#Update StandInit with constants
Query11 ='''UPDATE FVS_StandInit, Fvs_constants
            SET FVS_StandInit.Variant = [Fvs_constants].[Var],
            FVS_StandInit.Region = [Fvs_constants].[Reg],
            FVS_StandInit.Forest = [Fvs_constants].[For],
            FVS_StandInit.Brk_DBH = [Fvs_constants].[dbh_break],
            FVS_StandInit.DG_Trans = [Fvs_constants].[diam_transl],
            FVS_StandInit.DG_Measure = [Fvs_constants].[diam_span],
            FVS_StandInit.Site_Species = [Fvs_constants].[site_spp],
            FVS_StandInit.Site_Index = [Fvs_constants].[site_index];'''

cursor2.execute(Query11)
cnxn2.commit()


##########################################################




##########################################################

end = time.time()
tot_time = end - start

print "Total time elapsed is " + str(tot_time) + " seconds"

print "Finished"


