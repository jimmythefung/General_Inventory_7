# Retired


# Write a script that read from sold.csv and update number of sold items
# Description builder in HTML, write out model number and ID. This help inspecting title vs description

import csv
import os
import subprocess
import sys
import shutil
import datetime


def menu():
    bar = '='*70
    sList = [
        bar,
        'Excel Output Lite',
        '1 (print)   - Print Inventory',
        '2 (report)  - Generate offline reports and CSV listing files',
        '3 (sort)    - Choose how to sort Master Inventory',
        '4 (sold)    - Print sold and their locations',
        '5 (qty)     - Options to modify listing quantity',
        bar
        ]
    i=0
    while i < len(sList):
        print sList[i]
        i = i + 1
    choice = raw_input('Choose option: ')
    return choice.lower()

# Choice 1

def print_inventory(Prod_List):

    Prod_Dict = Dictionarize(Prod_List)

    # Ask user for brand input

    while True:
        choiceList = sorted(Prod_Dict)
        print 'Available brands:'
        for item in choiceList:
            print item
        print 'ALL'
            
        brand = raw_input('Brand name to print: ').upper()
        if (brand == 'ALL') or (brand in sorted(Prod_Dict)):
            break
        else:
            print 'Brand not found.'

    # Prepare save path

    savePath = 'c:\BLUESKY SCRIPTING\Printed Inventory\Output'
    l = 1
    while os.path.exists(savePath):
        savePath = 'c:\BLUESKY SCRIPTING\Printed Inventory\Output' + '_' + str(l)
        #print savePath
        l = l + 1
    os.makedirs(savePath)

    # Save to file
    
    heading = ['MAKER', 'MODEL', 'REV', 'QTY', 'LOCATION', 'QUANTITY UPDATE']
    Prod_Dict = Dictionarize(Prod_List)

    file_dst_ByModel = savePath + '\\' +brand+'_ByModel.csv'
    file_dst_ByLocation = savePath + '\\' +brand+'_ByLocation.csv'
    
    if brand != 'ALL':
        CustomSaveCSV(Dictionarize(Prod_Dict[brand]), heading, file_dst_ByModel)
        CustomSaveCSV(Dictionarize(Prod_Dict[brand], 'LOCATION'), heading, file_dst_ByLocation)
    else:
        for MAKER in sorted(Prod_Dict):
            file_dst_ByModel = savePath + '\\' +MAKER+'_ByModel.csv'
            file_dst_ByLocation = savePath + '\\' +MAKER+'_ByLocation.csv'
            CustomSaveCSV(Dictionarize(Prod_Dict[MAKER]), heading, file_dst_ByModel)
            CustomSaveCSV(Dictionarize(Prod_Dict[MAKER], 'LOCATION'), heading, file_dst_ByLocation)

    if brand == 'ALL':
        file_dst_ByModel = savePath + '\\' +'ALL'+'_ByModel.csv'
        file_dst_ByLocation = savePath + '\\' +'ALL'+'_ByLocation.csv'
        modelArr = printerFormat(Prod_List, heading, 'MODEL')
        locationArr = printerFormat(Prod_List, heading, 'LOCATION')
        CSV_ArrWrite(modelArr, file_dst_ByModel)
        CSV_ArrWrite(locationArr, file_dst_ByLocation)

        
        #file_dst_ByModel = savePath + '\\' +'ALL'+'_ByModel.csv'
        #file_dst_ByLocation = savePath + '\\' +'ALL'+'_ByLocation.csv'
        #CustomSaveCSV( Prod_Dict , heading, file_dst_ByModel)

        #Prod_Dict = Dictionarize(Prod_List, 'LOCATION')
        #CustomSaveCSV( Prod_Dict , heading, file_dst_ByLocation)
            
##    # Print to shell
##
##    if brand != 'ALL':
##        for item in Prod_Dict[brand]:
##            print item.data['MAKER'], item.data['MODEL'], 'Qty:'+item.data['QTY'], 'Price: $'+item.data['PRICE']
##        file_dst = savePath + '\\'+brand+'.csv'
##        saveInventory(Dictionarize(Prod_Dict[brand]), file_dst)
##    else:
##        for MAKER in sorted(Prod_Dict):
##            print ''
##            for item in Prod_Dict[MAKER]:
##                print item.data['MAKER'], item.data['MODEL'], 'Qty:'+item.data['QTY'], 'Price: $'+item.data['PRICE']
##            file_dst = savePath + '\\'+MAKER+'.csv'
##            saveInventory(Dictionarize( Prod_Dict[MAKER] ), file_dst)

    subprocess.Popen(r'explorer /select,'+ file_dst_ByModel)

# Takes prud_list and the order of prints (by location or model).
# return arrays of excel extries such that new brands start at multiples of 47th lines
def printerFormat(Prod_List, heading, sortOption):

    excelNewLine = 47
    #heading = Prod_List[0].dataNames
    heading = heading = ['MAKER', 'MODEL', 'REV', 'QTY', 'LOCATION', 'QUANTITY UPDATE']

    Prod_Dict = Dictionarize( Prod_List, sortOption )
    brandNames = sorted(Prod_Dict)

    collection = []

    for brand in brandNames: # brand = ['Delta', 'Hipro',...etc]

        vec47 = []
        vec47 = vec47 + [heading] # vec47 = [ [['PRODUCT'], ['MODEL'], ... etc] ] 
        brandArr = Prod_Dict[brand]

        # Keep building...
        for item in brandArr:
            if len(vec47) < excelNewLine:
                vec47 = vec47 + [item.customRow(heading)] #vec47 = [ [['PRODUCT'], ['MODEL'], ... etc], [['PSU'],['1956D'], ..etc], ... etc] 
            else:
                collection = collection + vec47
                vec47 = []
                vec47 = vec47 + [heading]
                vec47 = vec47 + [item.customRow(heading)]
                
        # Until at the end, fill up with empty excel lines
        while len(vec47) < excelNewLine:
            vec47 = vec47 + [['']]
        collection = collection + vec47

    #print len(collection)
    return collection

def CustomSaveCSV(Prod_Dict, heading, dst):
    
    fout = open(dst, 'wb')
    Fileout = csv.writer(fout)
    Fileout.writerow(heading)

    for MAKER in sorted(Prod_Dict):
        for item in Prod_Dict[MAKER]:
            Fileout.writerow(tuple(item.customRow(heading)))
            
    fout.close()

    
# Choice 2
        
def csv_report():
    #folder_check()
    
    # Initialize file paths and load our excel inventory
    Prod_List = load_data(inventoryFile)
    Prod_Dict = Dictionarize( Prod_List )
    
    offlineListPath = r'c:/BLUESKY SCRIPTING/Offline Listing/Offline Output'
    folder_index = 1
    while os.path.exists(offlineListPath):
        offlineListPath = 'c:\BLUESKY SCRIPTING\Offline Listing\Offline Output' + '_' + str(folder_index)
        folder_index = folder_index + 1
    os.makedirs(offlineListPath)

    offlineCSV = offlineListPath + '\\' + 'Report_offline.csv'

    print ''
    while True:

        choiceList = sorted(Prod_Dict)
        print 'Available brands:'
        for item in choiceList:
            print item
        print 'ALL'
        
        brand = raw_input('Brand name to check (Maker, not firm): ').upper()
        if (brand == 'ALL') or (brand in sorted(Prod_Dict)):
            break
        else:
            print 'Brand not found.'

    
    i = 1
    offline = 0
    online = 0
    heading = ['MAKER', 'MODEL', 'REV', 'QTY', 'LOCATION', 'STATUS', 'LISTING']
    #heading = Prod_List[0].dataNames
    #heading = ['TITLE', 'MODEL', 'PRICE', 'QTY', 'STATUS', 'LISTING']
    offline_list = [tuple(heading)]
    Prod_List = []
    offline_objects = []


    # Define what Prod_List is going to contain (depending on user choice)
    
    if brand != 'ALL':
        Prod_List = Prod_Dict[brand]
    else:
        for MAKER in sorted(Prod_Dict): #sorted gives arry of sorted keys
            Prod_List = Prod_List + Prod_Dict[MAKER] 



    # Create folder for offline objects. First collect all offline objects
    for item in Prod_List:
        if item.data['STATUS'] == 'OFFLINE':
            offline_objects = offline_objects + [item]
            offline = offline + 1
        else:
            online = online + 1


    # folder name: item.data['MAKER'] + '\\' + item.data['MODEL'] + ' ' + item.data['REV']
    for item in offline_objects:
        
        folder_name = item.data['MAKER'] + '\\' + item.data['MODEL'] + ' ' + item.data['REV']
        
        unlist_folder = offlineListPath + '\\'+'unlist'+'\\' + folder_name
        relist_folder = offlineListPath + '\\'+'relist'+'\\' + folder_name
        soldOut_folder = offlineListPath + '\\'+'sold out'+'\\' + folder_name

        if item.data['LISTING'] =='UNLIST':
            if not os.path.exists(unlist_folder):
                os.makedirs(unlist_folder)
        if item.data ['LISTING'] == 'RELIST':
            if not os.path.exists(relist_folder):
                os.makedirs(relist_folder)                
        if item.data ['LISTING'] == 'SOLD OUT':
            if not os.path.exists(soldOut_folder):
                os.makedirs(soldOut_folder)                



    # Write offline entries to excel file
    
    fout = open(offlineCSV, 'wb') # "wb" for overwrite, 'a' for appending single-spaced row entried
    
    FileOut = csv.writer(fout)

    unlist = [] # collects "UNLIST" products
    relist = [] # collects "RELIST" products
    soldOut = [] # collects "SOLD OUT" products

    for item in offline_objects:
        if item.data['LISTING'] == 'UNLIST':
            unlist = unlist + [item] 
        if item.data['LISTING'] == 'RELIST':
            relist = relist + [item]
        if item.data['LISTING'] == 'SOLD OUT':
            soldOut = soldOut + [item]

    FileOut.writerow(heading)

    for item in unlist:
        #FileOut.writerow(item.excelRow())
        FileOut.writerow(item.customRow(heading))

    for item in relist:
        #FileOut.writerow(item.excelRow())
        FileOut.writerow(item.customRow(heading))

    for item in soldOut:
        #FileOut.writerow(item.excelRow())
        FileOut.writerow(item.customRow(heading))

        
    print ''    
    print 'Number of offline: ' + str(offline)
    print 'Number of onlines: ' + str(online)
    print 'Total in inventory: ' + str(offline + online)
    print ''
 
    fout.close()

    # Generate a turbo lister file
    
    turboList( offlineListPath, offline_objects )

    # Generate a ebayUpload_revise.csv
    
    reviseArr = ebayFileOut(Prod_List, 'revise')
    CSV_ArrWrite(reviseArr, offlineListPath+'\\'+'ebayUpload_revise_'+str(folder_index)+'.csv')

    # Generate a ebayUpload_relist.csv
    
    relistArr = ebayFileOut(Prod_List, 'relist')
    CSV_ArrWrite(relistArr, offlineListPath+'\\'+'ebayUpload_relist_'+str(folder_index)+'.csv')

    # Generate a combined upload file

    combinedArr = reviseArr + relistArr[1:]
    CSV_ArrWrite(combinedArr, offlineListPath+'\\'+'ebayUpload_combined_'+str(folder_index)+'.csv')

    # Generate a report on quantity mismatch.

    subprocess.Popen(r'explorer /select,'+ offlineCSV)
    filePath = offlineListPath+'\\'+'Report_QuantityMismatch.csv'
    Revision_Report(Prod_List, filePath)
        

def CSV_ArrWrite(entry, filePath):
    
    fout = open(filePath, 'wb')
    FileOut = csv.writer(fout)
    
    for line in entry:
        FileOut.writerow(line)

    fout.close()
    
    
def extract_quantity(line):
        s1 = removeDSpace(line)
        s1 = s1.split()
        return s1[len(s1)-1]
                    

            
def turboList( offlineListPath, offline_List ): # NOTE! Different ebay selling category (template) applies!

    template_name = 'PSU.csv'
    template_path = 'c:/BLUESKY SCRIPTING/Listing Template/'+template_name
    
    # open turbo lister template
    
    fin = open(template_path)
    ReadFile = csv.reader(fin)

    # read template heading
    
    header = ReadFile.next() # header = ['MAKER', 'FIRM', ... ]

    # read template first entry
    
    temp_entry = ReadFile.next()
    fin.close()

    # Convert the template title-value pairs into a dictionary

    temp_dict = dict()
    j = 0
    while j < len(header):
        temp_dict[ header[j] ] = temp_entry[j]
        j += 1

    # Load offline entries
    
    #offlineCSV = offlineListPath + '/' + 'offline.csv'
    #offline_List =  load_data( offlineCSV )

    Unlist_Arr=[]
    Relist_Arr=[]
    SoldOut_Arr=[]

    for item in offline_List:
        
        # Collect Unlist items
        
        if item.data['LISTING'] == 'UNLIST':
            Unlist_Arr = Unlist_Arr + [item]

        # Collect Relist items
        
        if item.data['LISTING'] == 'RELIST':
            Relist_Arr = Relist_Arr + [item]

        # Collect SoldOut items

        if item.data['LISTING'] == 'SOLD OUT':
            SoldOut_Arr = SoldOut_Arr + [item]        

    outputArr = [Unlist_Arr, Relist_Arr, SoldOut_Arr]
    
    # Write template heading to output.csv
    
    Unlist = offlineListPath + '/' + 'turbo_Unlist.csv'
    SoldOut = offlineListPath + '/' + 'turbo_SoldOut.csv'
    Relist = offlineListPath + '/' + 'turbo_Relist.csv'
    turboFile = [Unlist, Relist, SoldOut]

    # For each output element
    i = 0
    while i < len(outputArr):
        fout = open(turboFile[i], 'wb')
        FileOut = csv.writer(fout)

        # Write header
        
        FileOut.writerow(header)

        # Revise template line for each product entry including:
        # Title, Quantity, Starting Price, Category 1

        for Prod_Obj in outputArr[i]:
            #temp_dict['Title'] = Prod_Obj.data['TITLE']
            temp_dict['Title'] = Prod_Obj.eBayTitle()
            temp_dict['Quantity'] = Prod_Obj.data['QTY']
            temp_dict['StartPrice'] = Prod_Obj.data['PRICE']
            temp_dict['Custom Label'] = Prod_Obj.data['LOCATION']

            # Construct entry line from the revised dictionary
            
            entry = []
            for title in header:
                entry = entry + [ temp_dict[title] ]

            # Write the built entry
            
            FileOut.writerow(entry)

        # Closing loop
        
        fout.close()
        i += 1
        
def getPath(filePath):
    pathNow = filePath
    i = 1
    while os.path.exists(filePath):
        pathNow = filePath+'_'+str(i)
        i = i + 1
    return pathNow
    
def saveInventory(Prod_Dict, saveFileLocationName):
    
    fout = open(saveFileLocationName, 'wb')
    for item in Prod_Dict:
        Prod_List = Prod_Dict[item]
        break
    heading = tuple(Prod_List[0].dataNames)
    Fileout = csv.writer(fout)
    Fileout.writerow(heading)

    # Prepare a line of empty entry of equal length to heading
    empty_list = ['']
    i=0
    while i < len(Prod_List[0].dataNames)-1:
        empty_list = empty_list + ['']
        i = i +1


    # Traverse through each brand or "MAKER"
    
    for MAKER in sorted(Prod_Dict): # Prod_Dict = { 'DELL':[...], ... }
        ProdArr = Prod_Dict[MAKER]  # ProdArr = [...]

        # each product object generates a line of excel entry with values correspond to heading
        for ProdObject in ProdArr:    # item is the PSU class object
            excel_entry = []
            i = 0
            while i < len(heading):
                excel_entry = excel_entry + [ProdObject.data[heading[i]]]
                i = i + 1
            Fileout.writerow(tuple(excel_entry))

        # At the end of writing all entries of a MAKER, add 10 empty lines
        counter = 0
        emptyLine_number = 20
        while counter < emptyLine_number:
            Fileout.writerow(tuple(empty_list))
            counter = counter + 1

    fout.close()
    #print 'Inventory Saved'

    
def saveBackup(src, dst, fileName_src, fileName_dst):

    srcFile = src + fileName_src + '.csv'
    dstFile = dst + fileName_dst

    if not os.path.exists(dst):
        os.makedirs(dst)

    i = 1
    while os.path.exists(dstFile + '.csv'):
        dstFile = dst + fileName_dst + '_' + str(i)
        i = i + 1

    shutil.copy2(srcFile, dstFile + '.csv')
        

def load_data(filepath):
    
    # Read files and load data into memory (Assign excel data to PSU object)

    fin = open(filepath)
    ReadFile = csv.reader(fin)

    header = ReadFile.next() # header = ['MAKER', 'FIRM', ... ]
    PSU_inventory=[]
    
    for row in ReadFile:
        #print row

        #initiate PSU class

        PSU_now = PSU()

        # Assign values to header elements
        
        i=0
        for item in row:
            #print i
            PSU_now.dataNames = header
            PSU_now.data[header[i]] = item
            #PSU_now.data[header[i]] = item.upper()
            i = i + 1


        # Collect into PSU_inventory

        PSU_inventory = PSU_inventory + [PSU_now]
        
    fin.close()

    return removeVoidEntries(PSU_inventory)

def folder_check():
    all_paths = [backupPath, templatePath, offlinePath, printPath, inputPath, outputPath]
    all_files = [active_ebayFile, sold_ebayFile, inventoryFile, templateFile]

    # Check folder existence, create them if not.
    i = 0
    while i < len(all_paths):
        print all_paths[i]
        if not os.path.exists(all_paths[i]):
            os.makedirs(all_paths[i])
        i = i + 1
            
    # Check essential file existence, prompt notice if does not exist (DNE).
    i = 0
    flag = 'ok'
    
    while i < len(all_files):
        if not os.path.exists(all_files[i]):
            print all_files[i] + ' [not found!]'
            flag = 'failed'
           #print "'"+all_files[i]+"' "+ 'not found. Put the said files there.'
        else:
            print all_files[i]
            #print "MASTER_INVENTORY.csv loaded"
        i = i + 1

    print ''
        
    if flag != 'ok':
        print 'Program aborted. Please make sure all essential files are in place.'
        sys.exit()
    else:
        print 'All files loaded sucessfully'
        print ''
            

class PSU:
    def __init__(self):
        self.dataNames=[]
        self.data = dict()

    # Construct ebay listing title
    def eBayTitle(self):
        prepString1 = removeDSpace( self.data['MAKER'] +' '+ self.data['POWER'] + 'W Power Supply '+ self.data['MODEL'] +' '+ self.data['REV']+' '+self.data['FIRM'] +' '+ self.data['ID1'] +' '+ self.data['ID2'] + ' Tested + Warranty' )
        prepString2 = removeDSpace( self.data['MAKER'] +' '+ self.data['POWER'] + 'W Power Supply '+ self.data['MODEL'] +' '+self.data['REV']+' '+self.data['FIRM'] +' '+ self.data['ID1'] +' '+ self.data['ID2'] + ' Tested + Warranty' )
        prepString3 = removeDSpace( self.data['MAKER'] +' '+ self.data['POWER'] + 'W Power Supply '+ self.data['MODEL'] +' '+self.data['REV']+' '+self.data['FIRM'] +' '+ self.data['ID1'] +' Tested + Warranty' )
        prepString4 = removeDSpace( self.data['MAKER'] +' '+ self.data['POWER'] + 'W Power Supply '+ self.data['MODEL'] +' '+self.data['REV']+' '+self.data['FIRM'] +' '+ self.data['ID1'] +' Tested + Wrty' )
        prepString5 = removeDSpace( self.data['MAKER'] +' '+ self.data['POWER'] + 'W Power Supply '+ self.data['MODEL'] +' '+self.data['REV']+' '+self.data['FIRM'] +' '+ self.data['ID1'])


        outStr = prepString1
        if len(outStr) > 80:
            outStr = prepString2
            if len(outStr) > 80:
                outStr = prepString3
                if len(outStr) > 80:
                    outStr = prepString4
                    if len(outStr) > 80:
                        outStr = prepString5
                        if (outStr) > 80:
                            print 'Title too long (>80): ' + outStr
                    
        return outStr
        
    # Gives the values to the column titles in an array
    def excelRow(self):
        output = []
        i = 0
        while i < len(self.dataNames):
            output = output + [self.data[self.dataNames[i]]]
            i = i + 1

        return output
    
    # Save as above, except the headings are customized
    
    def customRow(self, heading):
        
        # Collect heading that is contained in the Product object
        final_heading = []
        for item in heading:
            if item in self.data: #or (item == 'TITLE'):
                final_heading = final_heading + [item]
                #if item == 'TITLE':
                    #self.data['TITLE'] = self.eBayTitle()
                
        output = []
        i = 0
        while i < len(final_heading):
            output = output + [self.data[final_heading[i]]]
            i = i + 1

        return output


def brandNames(Prod_List):
    i = 0
    brandArr = []
    currentName = ''
    for item in Prod_List:
        if item.data['MAKER'] != currentName:
            currentName = item.data['MAKER']
            brandArr = brandArr + [currentName.upper()]
    brandArr.sort()
    return brandArr
            
def removeDSpace(s):
    s1 = s
    space = ' '
    space_s = ' '
    i = 2
    while i < 20:
        space_s = space* i
        while space_s in s1:
            s1 = s1.split(space_s)
            s1 = ' '.join(s1)
        i = i + 1
    return s1


def Dictionarize(Prod_List, sortOption = 'MODEL'):
    PartDict = dict()
    MakerList = brandNames(Prod_List)

    for maker in MakerList:
        temp = []
        for item in Prod_List:
            if maker == item.data['MAKER']:
                temp = temp + [item]
        PartDict[maker] = temp

    return DictSort( PartDict, sortOption)
                
def DictSort(Product_Dict, sortOption='MODEL' ): #To be used and sorted by PSU_Parting which builds a dictionary of array of PSU models by brand

    if sortOption == 'MODEL':
        for item in Product_Dict:
            Product_Dict[item] = sorted(Product_Dict[item], key=lambda PSU_Object: PSU_Object.data['MODEL']) #For each brand PSU, order by model number alpha-numerically
        return Product_Dict
            
    elif sortOption == 'LOCATION':
        sortOrder = sort_by_location()
        kept1 = []
        kept2 = []

        # Extracts items with assigned location letter
        
        for letter in sortOrder:
            for brand in Product_Dict:
                for item in Product_Dict[brand]:
                    if item.data['LOCATION'] == letter:
                        kept1 = kept1 + [item]
                        
        # Collect non-location-assigned items
        for brand in Product_Dict:
            for item in Product_Dict[brand]:
                if item.data['LOCATION'] not in sortOrder:
                    kept2 = kept2 + [item]
                    
        kept = kept1 + kept2
        keptDict = Dictionarize(kept, 'NONE')
        return keptDict

    else:
        return Product_Dict

def test(Prod_List):

    Prod_Dict = Dictionarize(Prod_List, 'LOCATION')
    #print Prod_Dict
    for brand in Prod_Dict:
        for item in Prod_Dict[brand]:
            print item.eBayTitle(), item.data['LOCATION']
    
def ListSort(Product_List):
    return sorted(Product_List, key=lambda item: item.data['Model'])

def sort_by_location():
    s = []
    mylist = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    mystring = ''
    for letter in mylist:
        i = 1
        while i < 30:
            s = s + [letter + str(i)]
            i = i + 1
    return tuple(s)

                               
def removeVoidEntries(Prod_List):
    keep = []
    for item in Prod_List:
        i = 0
        true = 0
        while i < len(item.dataNames):
            if item.data[item.dataNames[i]] == '':
                true = true + 1
            i = i + 1


        if true != i: #all excel column is not all empty
            keep = keep + [item]
        
            
    return keep


def Prod_Search(item, ebay_List):
    currentModel = item.data['MODEL'] +' ' # appends a space to Model to resolve a comparison bug with ebay.txt lines
    #item.data['TITLE'] = item.eBayTitle()

    #fin = open(ebayFilePath)
    qty = 0
                
    for ebayItem in ebay_List:
        line = ebayItem.data['Item Title']

        # Different brand requires slightly different comparison scheme
        # i.e.: Bestec; comparison requires both MODEL and HP part number, since a model may have multiple part number

        # Bestec
        if item.data['MAKER'] == 'BESTEC':
            if (currentModel in line) and ( item.data['ID1'] in line ) and (item.data['REV'] in line):
                flag = 'found'
                #qty = ebayItem.data['Quantity Available']
                break
            else:
                flag = 'unfound'

        # In general
        else:         
            if currentModel in line:
                flag = 'found'
                #qty = ebayItem.data['Quantity Available']
                break
            else:
                flag = 'unfound'

    
    return flag, ebayItem

def compare_ebay(Prod_List, active_List, sold_List):
    

    # Extract active listing information; Saerch for title in active_List
    for item in Prod_List:
        item.data['ebayID'] = ''
        item.data['ebayTitle'] = ''
        item.data['EBAY QTY'] = ''
        item.data['STATUS'] = ''
        item.data['LISTING'] = ''


        # Search in active.csv for online status
        
        flag, activeItem = Prod_Search(item, active_List)
        if flag == 'found':
            item.data['EBAY QTY'] = activeItem.data['Quantity Available']
            item.data['ebayID'] = 'Online: ' + activeItem.data['Item ID']
            item.data['ebayTitle'] = activeItem.data['Item Title']
            item.data['STATUS'] = 'O'
            item.data['LISTING'] = 'LISTED'


        # Seach in sold.csv for sold history. Note this search only applied to item not found in active.csv (hence the else)
        else:
            item.data['STATUS'] = 'OFFLINE'
            item.data['EBAY QTY'] = '0'
            flag, soldItem = Prod_Search(item, sold_List)

            if flag == 'found':
                item.data['ebayID'] = 'OFFLINE: ' + soldItem.data['Item ID']
                item.data['ebayTitle'] = soldItem.data['Item Title']
                item.data['LISTING'] = 'RELIST'
            else:
                item.data['LISTING'] = 'UNLIST'

        # Finally regardless of status, sold out is sold out
        if item.data['QTY'] == '0':
            item.data['LISTING'] = 'SOLD OUT'
            
    return Prod_List

            #This is to catch weird case when active contains sold item that has ebayqty = 0
            # Update: this bug does not exists; it happens with active.csv is update but sold.csv is outdated

##        # If not found, assign STATUS to OFFLINE, then assign LISTING TO: UNLIST, RELIST, or SOLD OUT
##        if flag == 'unfound':
##        
##            
####            if item.data['QTY'] == '0':
####                item.data['LISTING'] = 'SOLD OUT'
####            else:
##            if (item.data['STATUS'] == 'O') or (item.data['LISTING'] == 'LISTED'): #or (item.data['LISTING'] == 'SOLD OUT'): # This means the product had once been listed
##                item.data['LISTING'] = 'RELIST'                                         # Then it only needs to be relisted from record.
##            if (item.data['LISTING']) == '' or (item.data['LISTING'] == 'UNLIST'):
##                item.data['LISTING'] = 'UNLIST'
##
##            item.data['STATUS'] = 'OFFLINE'
##            item.data['EBAY QTY'] = '0'
##
##            # Search in sold_List to obtain ebay ID and title info
##
##            flag, soldItem = Prod_Search(item, sold_List)
##            
####            errorItem = []
####            if flag == 'unfound':
####                # For now just store them in errorItem array and do nothing to them.
####                # If not found, it be must unlisted.
####                errorItem = errorItem + [item]
####                item.data['LISTING'] = 'UNLIST'
####                #item.data['LISTING'] = 'UNLIST'
####                print item.eBayTitle() + ' not found in sold.csv (for offline entries). Assigned as UNLIST.'
####                #print soldItem.data['Item Title'] + ' not found in sold.csv (for offline entries). Assigned as UNLIST.'
####                #print len(errorItem)
##
##                
##            if flag == 'found': # if found in sold.csv, it needs to relist
##                item.data['LISTING'] = 'RELIST'
##                item.data['ebayID'] = 'Offline: ' + soldItem.data['Item ID']
##                item.data['ebayTitle'] = soldItem.data['Item Title']
##        if item.data['QTY'] == 0:
##            item.data['LISTING'] == 'SOLD OUT'
                
                

    return Prod_List

# Generate a csv file reporting items with mismatched qty, and a csv file for ebay upload to revise items.
def Revision_Report(Prod_List, ReportPath):

    # Part 1
    # Pick out the product with mismatched qty
    mismatch = []
    for item in Prod_List:
        if item.data['EBAY QTY'] != item.data['QTY']:
            ebayQty = item.data['EBAY QTY']
            #item.data['EBAY QTY'] = '**' + ebayQty
            mismatch = mismatch +[item]
            
    # Write the Mismatch prod objects to csv
    heading = ['MAKER', 'MODEL', 'REV', 'PRICE', 'QTY', 'EBAY QTY']

    CustomSaveCSV(Dictionarize(mismatch, 'NONE'), heading, ReportPath) # NONE means don't change the order of products


# Currently revise Qty and Price. Return a list of product whose Qty mismatch
def ebayFileOut(Prod_List, option):

    active_List = load_data(active_ebayFile)
    sold_List = load_data(sold_ebayFile)
    
    template_name = 'PSU.csv'
    template_path = 'c:/BLUESKY SCRIPTING/Listing Template/'+template_name
    
    # open turbo lister template
    
    fin = open(template_path)
    ReadFile = csv.reader(fin)

    # read template heading
    
    header = ReadFile.next() # header = ['MAKER', 'FIRM', ... ]

    # read template first entry
    
    temp_entry = ReadFile.next()
    fin.close()


    # Start to build output file now

    #reviseFile = 'c:\BLUESKY SCRIPTING\Output Files\\revise_upload.csv'
    #fout = open(reviseFile, 'wb')
    #FileOut = csv.writer(fout)
    #FileOut.writerow(header)

    output = []
    output = output + [header]

    if option == 'revise':
        for item in Prod_List:


            # Convert the template title-value pairs into a dictionary; all values are now empty string
            temp_dict = dict()
            j = 0
            while j < len(header):
                temp_dict[ header[j] ] = ''
                j += 1

            # Only the active items may be revised, so search in active_List
            flag, activeItem = Prod_Search(item, active_List)

            if flag == 'found' and item.data['QTY'] != activeItem.data['Quantity Available']:
                
                ## Modify the following code to revise Price and Quantity!!
                #temp_dict['Custom Label'] = item.data['LOCATION']
                if item.data['QTY'] == '0':
                    temp_dict['Action(CC=Cp1252)'] = 'End'
                else:
                    temp_dict['Action(CC=Cp1252)'] = 'Revise'
                #temp_dict['Title'] = activeItem.data['Item Title']
                temp_dict['ItemID'] = activeItem.data['Item ID']
                temp_dict['Quantity'] = item.data['QTY']
                temp_dict['StartPrice'] = item.data['PRICE']


                # Construct entry line from the revised dictionary
                
                entry = []
                for title in header:
                    entry = entry + [ temp_dict[title] ]

                # Add to output

                output = output + [entry]

                
    if option == 'relist': # Build output based on STATUS = RELIST      
        for item in Prod_List:

            # Search relistable item from sold listing
            flag, soldItem = Prod_Search(item, sold_List)
            
            # Convert the template title-value pairs into a dictionary; all values are now empty string
            temp_dict = dict()
            j = 0
            while j < len(header):
                temp_dict[ header[j] ] = ''
                j += 1

            # Applying the rule to build ebayUpload_relist.csv
            if flag == 'found' and (item.data['LISTING'] == 'RELIST'):

                ## Modify the following code to revise Title, Price and Quantity!!
                temp_dict['Action(CC=Cp1252)'] = 'Relist'
                temp_dict['Title'] = item.eBayTitle()
                temp_dict['Duration'] = 'GTC'
                temp_dict['ItemID'] = soldItem.data['Item ID']
                temp_dict['Quantity'] = item.data['QTY']
                temp_dict['StartPrice'] = item.data['PRICE']
                temp_dict['Custom Label'] = item.data['LOCATION']
                
                # Construct entry line from the revised dictionary
                entry = []
                for title in header:
                    entry = entry + [ temp_dict[title] ]

                # Add to output
                output = output + [entry]
                
            if flag == 'unfound' and (item.data['LISTING'] == 'RELIST'):
                print 'non-relistable item found while building ebayUpload_relist.csv'

            
    #FileOut.writerow(entry)
            
    #fout.close()

    return output
  
def reload_data(sortOption):
    print 'Soring by: ' + sortOption

    # Load master inventory
    Prod_List = load_data(inventoryFile)
    active_List = load_data(active_ebayFile)
    sold_List = load_data(sold_ebayFile)

    # Update inventory with online/offlinhe status and QTY by comparing with ebay.csv
    Prod_List = compare_ebay(Prod_List, active_List, sold_List)

    # Dictionarize, backup and save
    Prod_Dict = Dictionarize(Prod_List, sortOption)
    saveBackup(srcDir, backupDir, fileName, fileName)
    saveInventory(Prod_Dict, inventoryFile)

    return Prod_List, Prod_Dict

def set_sortBy_flag():
    global sort_by_flag
    print 'Sort MASTER INVENTORY by'
    print ' (blank space for NONE)'
    print '1 - Model'
    print '2 - Location'

    while True:
        answer = raw_input('Please enter one option above (Hit enter to default Model):').upper()
        if (answer == '') or (answer == '1') or (answer == 'MODEL') or (answer == '2') or (answer == 'LOCATION'):

            # Assign MODEL
            if (answer == '1') or (answer == 'MODEL'):
                sort_by_flag = 'MODEL'
            

            # Assign LOCATION
            if (answer == '2') or (answer == 'LOCATION'):
                sort_by_flag = 'LOCATION'

            if (answer == ''):
                sort_by_flag = 'NONE'

            break
                
        else:
            print 'Invalid input. Type the number (i.e. 1) or the name of sorting scheme (i.e. model).'
            pass

    subprocess.Popen(r'explorer /select,'+ inventoryFile)
##        
##def SalesUpdate(Prod_List, SalesFilePath):
##    for item in Prod_List:
##        item.data['SOLD'] = Obtain_SoldCount(item, SalesFilePath)
##    return Prod_List
##        
##def Obtain_SoldCount(item, SalesFilePath):
##    currentModel = item.data['MODEL'] + ' '
##    count = 0
##
##    fin = open(SalesFilePath)
##    
##    for line in fin:
##        if currentModel in line:
##            count = count + obtain_quantity(line)
##
##    fin.close()
##    
##    return count
def SalesProc(prod_List, file_sales):
    sales_List = load_data(file_sales)
    for item in Prod_List:
        
        flag, item = Prod_Search( sold, Prod_List )
        if flag == 'unfound':
            print sold.data['Item Title'] + ' not found!'
        else:
            line = sold.data['Item Title'] + ', Location: ' + item.data['LOCATION'] + ' QTY: ' + sold.data['Quantity']
            print line
    
def modQTY(Input_List, active_List, sold_List):

    Prod_List = Input_List[:]

    sList = [
        '',
        '1 - Synchronize QTY (i.e. perform option 2 and 3 below, then empty out the QTY UPDATE column)',
        '2 - Apply EBAY QTY to QTY',
        '3 - Add QTY UPDATE to QTY',
        '4 - Subtract recently sold from QTY (requires recentlySold.csv)',
        '',
        'You will be asked to confirm saving the changes again.',
        '',
        ]
    for s in sList:
        print s
    option = raw_input('Choose an option: ')

    if option == '1':
        for item in Prod_List:
            
            ebayQty = int(item.data['EBAY QTY'])
            if item.data['QTY UPDATE'].isdigit():
                updateQty = int(item.data['QTY UPDATE'])
            if item.data['QTY UPDATE'] == '':
                updateQty = 0
                
            print item.data['LOCATION'] + ' = ' + str(ebayQty) + ' + ' + str(updateQty)
            item.data['QTY'] = ebayQty + updateQty
            item.data['QTY UPDATE'] = ''

    elif option == '2':
        for item in Prod_List:
            item.data['QTY'] = item.data['EBAY QTY']

    elif option == '3':
        for item in Prod_List:
            if item.data['QTY UPDATE'].isdigit():
                print item.data['LOCATION'], ': ' + item.data['QTY'] + ' + ' + item.data['QTY UPDATE']
                item.data['QTY'] = int(item.data['QTY']) + int(item.data['QTY UPDATE'])
                print item.data['QTY']
                item.data['QTY UPDATE'] = ''

    elif option == '4':
        file_recentlySold = 'c:\BLUESKY SCRIPTING\Input Files\\recentlySold.csv'
        recentlySold_List = load_data(file_recentlySold)
        for soldItem in recentlySold_List:

            # loop through our Prod_list inventory
            for myItem in Prod_List:
                if len(myItem.data['ebayID'].split()) == 2 and myItem.data['ebayID'].split()[1] == soldItem.data['Item ID']:
                    print myItem.data['ebayTitle'], ': '+ str(myItem.data['QTY']) + ' - ' + str(soldItem.data['Quantity'])
                    #print myItem.data['LOCATION'], myItem.data['LISTING TITLE'], myItem.data['PRICE'], myItem.data['EBAY QTY'], ': '+ myItem.data['QTY'] + ' - ' + soldItem.data['Quantity']
                    myItem.data['QTY'] = str(int(myItem.data['QTY']) - int(soldItem.data['Quantity']))
                    print myItem.data['QTY']
                    if int(myItem.data['QTY']) < 0:
                        myItem.data['QTY'] = 0
                        print 'void (-) values to 0.'
                        

    else:
        pass

    confirm = raw_input('Save modifications to Master Template.csv? (y/n): ')
    if confirm == 'y':
        #saveItemList(file_inventory, Prod_List, 10)
        #Prod_Dict = Dictionarize(Prod_List, sortOption)
        #saveInventory(Prod_Dict, inventoryFile)
        print 'Modifications saved.'
        return Prod_List
    else:
        print 'Changes aborted. Returning to menu.'
        return Input_List    
####################
# Start of program
####################
 # Set default sort flag here
sort_by_flag = 'NONE' #Either: MODEL, LOCATION, or NONE


# Define paths and files
#Paths
backupPath = 'c:\BLUESKY SCRIPTING\Backup'
templatePath = 'c:\BLUESKY SCRIPTING\Listing Template'
offlinePath = 'c:\BLUESKY SCRIPTING\Offline Listing'
printPath = 'c:\BLUESKY SCRIPTING\Printed Inventory'
inputPath = 'c:\BLUESKY SCRIPTING\Input Files'
outputPath = 'c:\BLUESKY SCRIPTING\Output Files'

#Files
active_ebayFile = 'c:\BLUESKY SCRIPTING\Input Files\\active.csv'
sold_ebayFile = 'c:\BLUESKY SCRIPTING\Input Files\\sold.csv'
inventoryFile = 'c:\BLUESKY SCRIPTING\Input Files\\MASTER_INVENTORY.csv'
templateFile = 'c:\BLUESKY SCRIPTING\Listing Template\\PSU.csv'
file_sales = 'c:\BLUESKY SCRIPTING\\SalesHistory.csv'

folder_check()



# Save a backup of MASTER_INVENTORY

srcDir = 'c:\BLUESKY SCRIPTING\Input Files\\'
backupDir = 'c:\BLUESKY SCRIPTING\Backup\By Comparison\\'
fileName = 'MASTER_INVENTORY'
#saveBackup(srcDir, backupDir, fileName, fileName)



# Initilize by loading master inventory and update against ebay.txt
Prod_List, Prod_Dict = reload_data(sort_by_flag)

choice = ""
while choice != '7' or choice != 'ext':
#Note, all options must take care to save to inventory on its own if they happen to temper with the Prod_List

    # Get user input for choice
    choice = menu()
    if choice == '1' or choice == 'print':
        print_inventory(Prod_List)
        #dummy1 = raw_input('Hit enter to return to menu')
        
    elif choice == '2' or choice == 'report':
        csv_report()
        #dummy2 = raw_input('Hit enter to return to menu')
        
    elif choice == '3' or choice == 'sort':
        set_sortBy_flag()
        #dummy3 = raw_input('Hit enter to return to menu')
        
    elif choice == '4' or choice == 'sold':
        SalesProc(Prod_List, file_sales)
        #dummy4 = raw_input('Hit enter to return to menu')
        
    elif choice == '5' or choice == 'QTY':
        # Load files
        active_List = load_data(active_ebayFile)
        sold_List = load_data(sold_ebayFile)
        Prod_List = modQTY(Prod_List, active_List, sold_List)
        
        # Dictionarize, backup and save
        Prod_Dict = Dictionarize(Prod_List, 'MODEL')
        #saveBackup(srcDir, backupDir, fileName, fileName)
        saveInventory(Prod_Dict, inventoryFile)
        #dummy4 = raw_input('Hit enter to return to menu')
        
    elif choice == '6' or choice == 'exp':
        export_entry()
        #dummy4 = raw_input('Hit enter to return to menu')

    print ''
    Prod_List, Prod_Dict = reload_data(sort_by_flag)
