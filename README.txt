**This script requires the PyODBC module
This is a python script that converts a NED-2 database into a FVS ready input database. 
The script collects the necesary data through SQL queries and inserts it into the appropriate columns in the FVS db.
Plot level info is transferred directly from the NED-2 database. Some stand level data is transferred, while other fields are populated via a lookup table pre-made in the Blank_DB_copy
