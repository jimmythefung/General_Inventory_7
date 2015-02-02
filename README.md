# General-Inventory-Program-for-eBay
An inventory management solution for small businesses designed to work with eBay fileExchange platform for bulk listing directly, or through eBay's turbolister.

# What it does?
The Master Template.csv is used to keep track of your product inventory. Modify it manually to add new products.  
The main.py program loads Master Template.csv and provides user with an interface to perform various tasks:  
-Look up product in your inventory  
-Generate printable report of your inventory  
-Generate overview file of what is offline and any detected new product that neeeds to be listed on eBay  
-Generate eBay's turbolister or fileExchange compatitible files for bulk listing upload  
-Offline-Synchronization inventory quantity against active.csv  
-Update your inventory file, Master Template.csv with quantity and listing statuses by checking against active.csv and sold.csv  


# Requirements
1. Program files: main.py, module.py
2. ebay fileExchange participant
3. Most up-to-date active listing file requested from and downloaded from ebay fileExchange saved as active.csv
4. Most up-to-date sold listing file requested from and downloaded from ebay fileExchange saved as sold.csv
5. Master Template.csv: Modify it to add your new product.

# Sample demonstration usage
1. Run main.py once to generate needed directories
2. Download the input files active.csv, Master Template.csv, and sold.csv
3. Save the above 3 files in your C:\BLUESKY SCRIPTING\Input Files\
4. Close all files and windows. Re-run the script main.py.
5. Follow on-screen instruction to interact with inventory 
