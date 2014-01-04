#Declare warning when overlapped title found; maybe dont pull ebayTitle in such cases. Use false count in addition to matchCount
#Check why item is sold but dont have record of ebay iD
#Ask user input for sort instruction
#Improve search algothim by spliting all ID, count all match and mismatch
#Generate temlpate by identifying ebay product category ID. Have it pass as to csv report.
#hard code header of template file

# Bug in power supply relist file

#View ebay categor
#http://listings.ebay.com/_W0QQloctZShowCatIdsQQsocmdZListingCategoryList


import csv
import os
import subprocess
import sys
import shutil
import datetime


ebayCatNum = dict()
ebayCatNum['POWER SUPPLY']    = 42017
ebayCatNum['MOTHERBOARD']     = 1244
ebayCatNum['VIDEO CARD']      = 175673
ebayCatNum['CAR PART']        = 6001
ebayCatNum['HARD DRIVE']      = 165

template_header = ['Action(CC=Cp1252)', 'SiteID', 'Format', 'Title', 'Condition', 'ConditionDescription', 'SubTitle', 'Custom Label', 'Category', 'Category2', 'StoreCategory', 'StoreCategory2', 'Quantity', 'LotSize', 'Currency', 'StartPrice', 'BuyItNowPrice', 'ReservePrice', 'InsuranceOption', 'InsuranceFee', 'DomesticInsuranceOption', 'DomesticInsuranceFee', 'PackagingHandlingCosts', 'InternationalPackagingHandlingCosts', 'Duration', 'PrivateAuction', 'Country', 'ProductIDType', 'ProductIDValue', 'Product:ProductReferenceID', 'ItemID', 'Description', 'HitCounter', 'PicURL', 'BoldTitle', 'Featured', 'GalleryType', 'FeaturedFirstDuration', 'Highlight', 'Border', 'HomePageFeatured', 'Subtitle in search resutls', 'GiftIcon', 'GiftServices-1', 'GiftServices-2', 'GiftServices-3', 'SalesTaxPercent', 'SalesTaxState', 'ShippingInTax', 'UseTaxTable', 'PostalCode', 'ProxyItem', 'VATPercent', 'Location', 'ImmediatePayRequired', 'PayPalAccepted', 'PayPalEmailAddress', 'PaymentInstructions', 'PaymateAccepted', 'ProPayAccepted', 'MoneyBookersAccepted', 'StandardPayment', 'UPC', 'CCAccepted', 'AmEx', 'Discover', 'VisaMastercard', 'IntegratedMerchantCreditCard', 'COD', 'CODPrePayDelivery', 'PostalTransfer', 'MOCashiers', 'PersonalCheck', 'MoneyXferAccepted', 'MoneyXferAcceptedinCheckout', 'PaymentOther', 'OtherOnlinePayments', 'PaymentSeeDescription', 'Escrow', 'ShippingType', 'GlobalShippingService', 'ShipFromZipCode', 'ShippingIrregular', 'ShippingPackage', 'WeightMajor', 'WeightMinor', 'WeightUnit', 'MeasurementUnit', 'ShippingDetails/CODCost', 'PackageLength', 'PackageWidth', 'PackageDepth', 'DomesticRateTable', 'InternationalRateTable', 'CharityID', 'CharityName', 'DonationPercent', 'ShippingService-1:Option', 'ShippingService-1:Cost', 'ShippingService-1:AdditionalCost', 'ShippingService-1:Priority', 'ShippingService-1:FreeShipping', 'ShippingService-1:ShippingSurcharge', 'ShippingService-2:Option', 'ShippingService-2:Cost', 'ShippingService-2:AdditionalCost', 'ShippingService-2:Priority', 'ShippingService-2:ShippingSurcharge', 'ShippingService-3:Option', 'ShippingService-3:Cost', 'ShippingService-3:AdditionalCost', 'ShippingService-3:Priority', 'ShippingService-3:ShippingSurcharge', 'ShippingService-4:Option', 'ShippingService-4:Cost', 'ShippingService-4:AdditionalCost', 'ShippingService-4:Priority', 'ShippingService-4:ShippingSurcharge', 'ShippingService-5:Option', 'ShippingService-5:Cost', 'ShippingService-5:AdditionalCost', 'ShippingService-5:Priority', 'ShippingService-5:ShippingSurcharge', 'GetItFast', 'DispatchTimeMax', 'IntlShippingService-1:Option', 'IntlShippingService-1:Cost', 'IntlShippingService-1:AdditionalCost', 'IntlShippingService-1:Locations', 'IntlShippingService-1:Priority', 'IntlShippingService-2:Option', 'IntlShippingService-2:Cost', 'IntlShippingService-2:AdditionalCost', 'IntlShippingService-2:Locations', 'IntlShippingService-2:Priority', 'IntlShippingService-3:Option', 'IntlShippingService-3:Cost', 'IntlShippingService-3:AdditionalCost', 'IntlShippingService-3:Locations', 'IntlShippingService-3:Priority', 'IntlShippingService-4:Option', 'IntlShippingService-4:Cost', 'IntlShippingService-4:AdditionalCost', 'IntlShippingService-4:Locations', 'IntlShippingService-4:Priority', 'IntlShippingService-5:Option', 'IntlShippingService-5:Cost', 'IntlShippingService-5:AdditionalCost', 'IntlShippingService-5:Locations', 'IntlShippingService-5:Priority', 'IntlAddnlShiptoLocations', 'PaisaPayAccepted', 'PaisaPay EMI payment', 'BasicUpgradePackBundle', 'ValuePackBundle', 'ProPackPlusBundle', 'BestOfferEnabled', 'AutoAccept', 'BestOfferAutoAcceptPrice', 'AutoDecline', 'MinimumBestOfferPrice', 'BestOfferRejectMessage', 'LocalOnlyChk', 'LocalListingDistance', 'BuyerRequirements:ShipToRegCountry', 'BuyerRequirements:ZeroFeedbackScore', 'BuyerRequirements:MinFeedbackScore', 'BuyerRequirements:MaxUnpaidItemsCount', 'BuyerRequirements:MaxUnpaidItemsPeriod', 'BuyerRequirements:MaxItemCount', 'BuyerRequirements:MaxItemMinFeedback', 'BuyerRequirements:LinkedPayPalAccount', 'BuyerRequirements:VerifiedUser', 'BuyerRequirements:VerifiedUserScore', 'BuyerRequirements:MaxViolationCount', 'BuyerRequirements:MaxViolationPeriod', 'SellerDetails:PrimaryPhone', 'SellerDetails:SecondaryPhone', 'ExtSellerDetails:Hours1Days', 'ExtSellerDetails:Hours1AnyTime', 'ExtSellerDetails:Hours1From', 'ExtSellerDetails:Hours1To', 'ExtSellerDetails:Hours2Days', 'ExtSellerDetails:Hours2AnyTime', 'ExtSellerDetails:Hours2From', 'ExtSellerDetails:Hours2To', 'ExtSellerDetails:TimeZoneID', 'ListingDesigner:LayoutID', 'ListingDesigner:ThemeID', 'ShippingDiscountProfileID', 'InternationalShippingDiscountProfileID', 'Apply Profile Domestic', 'Apply Profile International', 'PromoteCBT', 'ShipToLocations', 'CustomLabel', 'CashOnPickup', 'ReturnsAcceptedOption', 'ReturnsWithinOption', 'RefundOption', 'ShippingCostPaidBy', 'ReturnsRestockingFee', 'WarrantyOffered', 'WarrantyType', 'WarrantyDuration', 'AdditionalDetails', 'MarketplaceType', 'ProjectGoodCategory', 'ShortDescription', 'ProducerDescription', 'RegionOfOrigin', 'ProducerPhotoURL', 'C:Wattage', 'C:Brand', 'C:Max. Output Power', 'C:Cooling', 'C:Modular', 'C:Form Factor', 'C:MPN', 'C:Connectors', 'C:Connection Type', 'C:Power Supply Size', 'C:Type', 'C:Model', 'C:Country of Manufacture', 'C:Title', 'C:Model Number', 'C:Compatibility', 'C:Weight', 'Relationship', 'RelationshipDetails']
template_values = ['Add', 'US', 'FixedPriceItem', 'PSU TITLE', '3000', '', '', '', '42017', '', '1', '0', '1', '', 'USD', '24.99', '', '', '', '', '', '', '', '', 'GTC', '0', 'US', '', '', '', '', 'Tested working. <BR>Guaranteed against DOA. <BR>30 days warranty. <BR>Continental US buyers only. <BR>Please check my listing for more computer parts.\n<P>\n<CENTER><A href="http://pages.ebay.com/selling_manager_pro"><IMG border=0 src="http://pics.ebay.com/aw/pics/sellingmanager/sellingmanagerPro/smPro_248x50.gif" width=248 height=50></A></CENTER>\n<P></P>\n\t\t\t', '', '', '0', '0', 'Gallery', '', '0', '0', '0', '0', '0', '', '', '', '', '', '', '0', '27529', '', '', '', '0', '1', 'wei@nc.rr.com', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '0', '', 'Flat', '0', '', '', 'PackageThickEnvelope', '', '', '', 'English', '', '', '', '', '', '', '', '', '', 'Other', '0', '0', '1', '1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '0', '1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '0', '', '', '', '', '', '', '', '', '', '', '', '-1', '2', 'Days_30', '', '', '1', '', '', '4', 'Days_30', '', '', '', '', '', '', '', '', '', '', '', '10000', '10', '0||', '0||', '0', '0', '', '', '', '', 'ReturnsAccepted', 'Days_30', 'MoneyBack', 'Buyer', 'NoRestockingFee', '', '', '', 'Restocking fees: No', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']

indexedBrand = ['ASUS', 'BIOSTAR', 'COMPAQ', 'DELL', 'ECS', 'EMACHINES', 'GIGABYTE', 'GATEWAY', 'HP', 'IBM', 'INTEL', 'MSI']

# Essential Folders
backupPath = 'c:\BLUESKY SCRIPTING\Backup'
templatePath = 'c:\BLUESKY SCRIPTING\Listing Template'
offlinePath = 'c:\BLUESKY SCRIPTING\Offline Listing'
printPath = 'c:\BLUESKY SCRIPTING\Printed Inventory'
inputPath = 'c:\BLUESKY SCRIPTING\Input Files'
outputPath = 'c:\BLUESKY SCRIPTING\Output Files'
all_paths = [backupPath, templatePath, offlinePath, printPath, inputPath, outputPath]

# Input Files
file_active = 'c:\BLUESKY SCRIPTING\Input Files\\active.csv'
file_sold = 'c:\BLUESKY SCRIPTING\Input Files\\sold.csv'
file_inventory = 'c:\BLUESKY SCRIPTING\Input Files\\Master Template.csv'
#file_template = 'c:\BLUESKY SCRIPTING\Listing Template\\PSU.csv'
all_files = [file_active, file_sold, file_inventory] 

# Output paths
path_offline = 'c:\BLUESKY SCRIPTING\Offline Listing\Offline Output'


# Output Files
file_backup = 'c:\BLUESKY SCRIPTING\Backup\\Master Template.csv'





########################
# Function Definitions #
########################

##def test():
##    #Prod_List = load_data(file_inventory)
##    #print Prod_List[0].dataNames
##
##    template_List = load_data(templatePath+'\\'+'MOTHERBOARD.csv')
##    print template_List[0].dataNames
##    print template_List[0].dataValues
##
##    sys.exit()
    
# Choice 1
def print_inventory(Prod_List, printPath):
    
    printPath = getIndexedPath(printPath + '\\'+ 'Output')
    os.makedirs(printPath)
    productNames = obtain_variation( Prod_List, 'PRODUCT' )
    heading = ['PRODUCT', 'BRAND', 'LISTING TITLE', 'LOCATION', 'QTY UPDATE']
    #heading = Prod_List[0].flexibleNames + ['QTY', 'LOCATION', 'QTY UPDATE']
    
    for product in productNames: # create product folders

        prodNow = getSubList( Prod_List, product )
        
        brandNames = obtain_variation( prodNow, 'BRAND' )
        
        for brand in brandNames:

            # Create current dir: c:\BLUESKY SCRIPTING\Printed Output\Output_i \MOTHERBOARD \DELL
            currentDir = printPath + '\\' + product + '\\' + brand
            os.makedirs(currentDir)
            
            brandArr = []

            for item in prodNow:
                if item.data['BRAND'] == brand:
                    brandArr = brandArr + [item]

            # Save by model
            brand_byModel = printPath + '\\' + product + '\\' + brand + '\\' + brand +'_by_Model.csv'
            brandArr = listSort( brandArr, 'alpha' )
            saveItemList(brand_byModel, brandArr, heading)

            # Save by location
            brand_byLocation = printPath + '\\' + product + '\\' + brand + '\\' + brand +'_by_Location.csv'
            brandArr = listSort (brandArr, 'LOCATION')
            saveItemList(brand_byLocation, brandArr, heading)

            
        # Save a outter scope file by model and location as well
        prodDir = printPath + '\\' + product + '\\' + product + '_by_Model.csv'
        alphaSorted_Prod = listSort( prodNow, 'alpha' )
        byModelArr = printerFormat_General(alphaSorted_Prod, heading)
        CSV_ArrWrite(byModelArr, prodDir)

        prodDir = printPath + '\\' + product + '\\' + product + '_by_Location.csv'
        locationSorted_Prod = listSort( prodNow, 'location' )
        byLocationArr = printerFormat_General(locationSorted_Prod, heading)
        CSV_ArrWrite(byLocationArr, prodDir)
        
    
    subprocess.Popen(r'explorer /select,'+ printPath)



# Choice 2
def csv_report(Prod_List, active_List, sold_List, path_offline, templatePath):


    # obtain folder path then make folder   
    path_offline = getIndexedPath( path_offline )
    makePath( path_offline )


    # Define what Prod_List contains depending on user input


    # Collect offline items
    online = 0
    offline = 0
    offline_objects = []
    for item in Prod_List:
        if item.data['STATUS'] == 'OFFLINE':
            offline_objects = offline_objects + [item]
            offline = offline + 1
        else:
            online = online + 1

    
    if len(offline_objects) == 0:
        print ''
        print 'Congrats! There are no offline listing! (If this msg is errorneous, check excel for any typos.'
        print ''
        s = raw_input('Press any key to exit.')
        print 'Program exited.'
        sys.exit()

    # Create folders for offline objects
    for item in offline_objects:

        # folder name: c:\BLUESKY SCRIPTING\Offline Listing\Offline_Report.csv\Motherboard\Dell\ebay title
        #folder_name = item.data['PRODUCT'] + '\\' + item.data['BRAND'] + '\\' + item.eBayTitle()

        # create folder name free of ":" and '.'
        s = item.data['LISTING TITLE']
        s = list(s)
        i = 0
        while i < len(s):
            if s[i] == '.' or s[i] == ':' or s[i] == '/' or s[i] == '\\':
                s[i] = ' '
            if s[i] == '/' or s[i] == '\\':
                s[i] = ''
                
            i = i + 1
        s = "".join(s)

        sOut = item.data['LOCATION'] + ' ' + s
        
        folder_name = removeDSpace(sOut)
        unlist_folder = path_offline + '\\' + item.data['PRODUCT'] +'\\'+'unlist'+'\\' + folder_name
        relist_folder = path_offline + '\\' + item.data['PRODUCT'] + '\\'+'relist'+'\\' + folder_name
        soldOut_folder = path_offline + '\\' + item.data['PRODUCT'] + '\\'+'sold out'+'\\' + folder_name
        threeFolders = [unlist_folder, relist_folder, soldOut_folder]

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
    # Improve algorithm here; takes a lot of time.
    productNames = obtain_variation(offline_objects, 'PRODUCT')

    for product in productNames:
        
        heading = ['LISTING TITLE', 'QTY', 'PRICE', 'LOCATION','LISTING']
        offlineCSV = path_offline + '\\'+ product+'\\' + product +'_Offline_Report.csv'
        fout = open(offlineCSV, 'wb') # "wb" for overwrite, 'a' for appending single-spaced row entried
        FileOut = csv.writer(fout)
        unlist = [] # collects "UNLIST" products
        relist = [] # collects "RELIST" products
        soldOut = [] # collects "SOLD OUT" products

        for item in offline_objects:
            if item.data['PRODUCT'] == product and item.data['LISTING'] == 'UNLIST':
                unlist = unlist + [item] 
            if item.data['PRODUCT'] == product and item.data['LISTING'] == 'RELIST':
                relist = relist + [item]
            if item.data['PRODUCT'] == product and item.data['LISTING'] == 'SOLD OUT':
                soldOut = soldOut + [item]

        FileOut.writerow(heading)

        for item in unlist:
            FileOut.writerow(item.customRow(heading))
            
        for item in relist:
            FileOut.writerow(item.customRow(heading))
            
        for item in soldOut:
            FileOut.writerow(item.customRow(heading))

        fout.close()
        savePath = path_offline + '\\' + product

# Turbo Lister
        
        # Generate a turbo lister file, ebayUpload_revise.csv, ebayUpload_relist.csv, and ebayUpload_all.csv
        turboList( offline_objects, path_offline, templatePath )
        
        # Generate a ebayUpload_revise.csv; getSubList(Prod_List, product) contains all items under motherboard then power supply in next loop
        reviseArr = ebayFileOut( getSubList(Prod_List, product), active_List, sold_List, path_offline, templatePath, 'revise')
        CSV_ArrWrite(reviseArr, savePath+'\\'+'ebayUpload_revise.csv')

        # Generate a ebayUpload_relist.csv
        relistArr = ebayFileOut( getSubList(Prod_List, product), active_List, sold_List, path_offline, templatePath, 'relist')
        CSV_ArrWrite(relistArr, savePath+'\\'+'ebayUpload_relist.csv')

# ebay Listing

        # Generate a ebayUpload_all.csv
        allArr = reviseArr + relistArr[1:]
        CSV_ArrWrite(allArr, savePath+'\\'+'ebayUpload_all.csv')

    #print path_offline+'\\'+ product    
    subprocess.Popen(r'explorer /select,'+ path_offline+'\\'+ product)




def ebayFileOut(Prod_List, active_List, sold_List, path_offline, templatePath, option):

    header = template_header
    # remove the space in template's custom label to work with ebayUpload file exchange format
    i = 0
    while i < len(header):
        if header[i] == 'Custom Label':
            header[i] = 'CustomLabel'
        i = i + 1
            
    temp_entry = template_values


    # Convert the template title-value pairs into a dictionary
    temp_dict = dict()
    j = 0
    while j < len(header):
        temp_dict[ header[j] ] = temp_entry[j]
        j += 1


    # Start to build output file now
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
            
            # Revise if:
            # 1. Item is online
            # 2. Item's LOCATION is not the Custom Label
            # 3. QTY mismatch
            # 4. ebayID mismatch against LISTING IDindexedBrand
            if flag == 'found' and (item.data['QTY'] != activeItem.data['Quantity Available']): #or (flag == 'found' and (item.data['SOLD'] == '0') and (item.data['LISTING TITLE'] != activeItem.data['Item Title'])): 
                #print flag
                
                ## Modify the following code to revise Price and Quantity!!
                if item.data['QTY'] == '0':
                    temp_dict['Action(CC=Cp1252)'] = 'End'
                else:
                    temp_dict['Action(CC=Cp1252)'] = 'Revise'

##                if item.data['SOLD'] == 0:
##                    temp_dict['Title'] = item.data['LISTING TITLE']

                temp_dict['ItemID'] = item.data['LISTING ID'].split()[1]
                temp_dict['Quantity'] = item.data['QTY']
                temp_dict['StartPrice'] = item.data['PRICE']
                temp_dict['Category'] = item.data['ebayCategory']
                temp_dict['CustomLabel'] = item.data['LOCATION']


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
            #if flag == 'found' and (item.data['LISTING'] == 'RELIST'):
            if (item.data['LISTING'] == 'RELIST') and len(item.data['LISTING ID'].split()) == 2: #and str(item.data['LISTING ID'].split()[1]).isdigit():

                ## Modify the following code to revise Title, Price and Quantity!!
                temp_dict['Action(CC=Cp1252)'] = 'Relist'
                temp_dict['Title'] = item.data['LISTING TITLE']
                temp_dict['Duration'] = 'GTC'
                temp_dict['ItemID'] = item.data['LISTING ID'].split()[1]
                temp_dict['Quantity'] = item.data['QTY']
                temp_dict['StartPrice'] = item.data['PRICE']
                temp_dict['CustomLabel'] = item.data['LOCATION']
                
                # Construct entry line from the revised dictionary
                entry = []
                for title in header:
                    entry = entry + [ temp_dict[title] ]

                # Add to output
                output = output + [entry]

            else:
                if (item.data['LISTING'] == 'RELIST'):
                    print 'non-relistable item found while building ebayUpload_relist.csv'
                    print len(item.data['LISTING ID'].split())


    return output
        #reviseArr = output
        #savePath = productFolder+'\\'+'ebayUpload'+'_'+option+'.csv'
        #print productFolder+'\\'+'ebayUpload'+'_'+option+'.csv'
        #CSV_ArrWrite(reviseArr, savePath)

        


def turboList( offline_List, path_offline, templatePath ):
    # path_offline = c:\BLUESKY SCRIPTING\Offline Listing\Offline Output_i
    # templatePath = 'c:\BLUESKY SCRIPTING\Listing Template'

    # Load product template and create folders; i.e. MOTHERBOARD.csv, POWER SUPPLY.csv
    productNames = obtain_variation( offline_List, 'PRODUCT' )

    for product in productNames:
    
        header = template_header
        temp_entry = template_values

        #obtain the current product being manipulated (by product name)
        offline_now = []
        for item in offline_List:
            if item.data['PRODUCT'] == product:
                offline_now = offline_now + [item]

        # create product folder (update: already created outside function call)
        productFolder = path_offline + '\\' + product # c:\BLUESKY SCRIPTING\Offline Listing\Offline Output_i\MOTHERBOARD


        # Convert the template title-value pairs into a dictionary
        temp_dict = dict()
        j = 0
        while j < len(header):
            temp_dict[ header[j] ] = temp_entry[j]
            j += 1


        # Load offline entries
        Unlist_Arr=[]
        Relist_Arr=[]
        SoldOut_Arr=[]
        for item in offline_now:
            
            # Collect Unlist items            
            if item.data['LISTING'] == 'UNLIST':
                Unlist_Arr = Unlist_Arr + [item]

            # Collect Relist items            
            if item.data['LISTING'] == 'RELIST':
                Relist_Arr = Relist_Arr + [item]

            # Collect SoldOut items
            if item.data['LISTING'] == 'SOLD OUT':
                SoldOut_Arr = SoldOut_Arr + [item]        

        
        # Write template heading to output.csv
        
        Unlist = productFolder + '\\' + 'turbo_Unlist.csv'
        
        SoldOut = productFolder + '\\' + 'turbo_SoldOut.csv'
        
        Relist = productFolder + '\\' + 'turbo_Relist.csv'

        # Collect above into arrays
        outputArr = [Unlist_Arr, Relist_Arr, SoldOut_Arr]
        turboFile = [Unlist, Relist, SoldOut]
        #outputArr = [Unlist_Arr]
        #turboFile = [Unlist]

        # For each output element in outputArr, perform write to csv
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
                temp_dict['Title'] = Prod_Obj.data['LISTING TITLE']
                temp_dict['Quantity'] = Prod_Obj.data['QTY']
                temp_dict['StartPrice'] = Prod_Obj.data['PRICE']
                temp_dict['Category'] = Prod_Obj.data['ebayCategory']
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
        

def CSV_ArrWrite(entry, filePath):
    
    fout = open(filePath, 'wb')
    FileOut = csv.writer(fout)
    
    for line in entry:
        FileOut.writerow(line)

    fout.close()
    
def listSort(Prod_List, sortBy='alpha'):
    sortBy = sortBy.upper()

    if sortBy == 'NONE':
        return Prod_List
    
    prod_dict = Dictionarize(Prod_List, 'PRODUCT')
    keepProd = []
    
    for prod in sorted(prod_dict): # prod = MOBO, PSU
        keepBrand = []
        
        brand_dict = Dictionarize(prod_dict[prod], 'BRAND') #brand_dict = ['EMACHINES': [...], 'ASUS': [...], ...etc]

        for brand in sorted(brand_dict): # brand = ['ASUS'. 'BIOSTAR', 'DELL', ...etc]
            #make a copy of brand list
            brandList = brand_dict[brand]

            #sort the items by alphabetical order ebaytitle
            if sortBy == 'ALPHA':
                brandList = sorted(brandList, key=lambda PSU_Object: PSU_Object.data['LISTING TITLE'])
            

            # sort by location number
            elif sortBy == 'LOCATION':
                #brandList = sorted(brandList, key=lambda PSU_Object: PSU_Object.data['LOCATION'])
                brandList = locationSort(brandList)
                
            
            #overwrite the brand_dict
            #brand_dict[brand] = brandList
            keepBrand = keepBrand + brandList

        keepProd = keepProd + keepBrand
    return keepProd

def getSubList(Prod_List, product, brand='all'):
    
    brand = brand.upper()
    output = []
    
    if brand == 'ALL':
        for item in Prod_List:
            if item.data['PRODUCT'] == product:
                output = output + [item]
        return output
    
    else:
        for item in Prod_List:
            if (item.data['PRODUCT'] == product) and (item.data['BRAND'] == brand):
                output = output + [item]

        return output


def locationSort(brandList):

    brandNow = brandList[0].data['BRAND']
    
    # Build sort string
    s = []
    end = len(brandList) + 1
    i = 1
    while i < end:
        s = s + [brandNow + ' ' + str(i)]
        i = i + 1

    #sortOrder = sort_by_location()
    kept1 = []
    kept2 = []
    s = tuple(s)

    # Extracts items with assigned location letter
    for letter in s:
        for item in brandList:
            if item.data['LOCATION'] == letter:
                kept1 = kept1 + [item]

    # Collect non-location assigned items
    for item in brandList:
        if item.data['LOCATION'] not in s:
            kept2 = kept2 + [item]

    kept = kept1 + kept2
    return kept

##def locationSort(brandList):
##    
##    # build sort string
##    s = []
##    mylist = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
##    mystring = ''
##    for letter in mylist:
##        i = 1
##        while i < 30:
##            s = s + [letter + str(i)]
##            i = i + 1
##    #sortOrder = sort_by_location()
##    kept1 = []
##    kept2 = []
##    s = tuple(s)
##
##    # Extracts items with assigned location letter
##    for letter in s:
##        for item in brandList:
##            if item.data['LOCATION'] == letter:
##                kept1 = kept1 + [item]
##
##    # Collect non-location assigned items
##    for item in brandList:
##        if item.data['LOCATION'] not in s:
##            kept2 = kept2 + [item]
##
##    kept = kept1 + kept2
##    return kept

def saveItemList(filePath, Prod_List, spacing = 0, heading = 'default'):
    
    if heading == 'default':
        fout = open(filePath, 'wb')
        heading = tuple(Prod_List[0].dataNames)
        Fileout = csv.writer(fout)
        Fileout.writerow(heading)
        
        brandNow = Prod_List[0].data['BRAND']

        for item in Prod_List:

            # Write void entries (spacing)
            if item.data['BRAND'] != brandNow:
                brandNow = item.data['BRAND']
                i = 0
                while i < spacing:
                    Fileout.writerow(['']*len(item.dataNames))
                    i = i + 1
            
            Fileout.writerow(item.excelRow())
        fout.close()
        
    else:
        fout = open(filePath, 'wb')
        Fileout = csv.writer(fout)
        Fileout.writerow(heading)

        for item in Prod_List:
            Fileout.writerow(item.customRow(heading))
        fout.close()

def compare_ebay(Prod_List, active_List, sold_List):
    

    # Extract active listing information; Saerch for title in active_List
    for item in Prod_List:
        print item.data['LOCATION']
        item.data['ebayID'] = ''
        item.data['ebayTitle'] = ''
        item.data['EBAY QTY'] = ''
        item.data['STATUS'] = ''
        item.data['LISTING'] = ''
        if item.data['QTY'] == '':
            item.data['QTY'] = '0'
        #item.data['ebayCategory'] = ''
        #item.data['myTitle'] = removeDSpace(item.warrantedTitle())
        


        # Search in active.csv for online status
        
        flag, activeItem = Prod_Search(item, active_List)
        
        if flag == 'found':
            item.data['ebayCategory'] = activeItem.data['Category Number']
            item.data['EBAY QTY'] = activeItem.data['Quantity Available']
            item.data['ebayID'] = 'ID: '+ activeItem.data['Item ID']
            #item.data['ebayID'] = 'Online: ' + activeItem.data['Item ID']
            
            #if item.data['LISTING ID'] != item.data['ebayID'].split()[1]: # May be superflorus statement; actually this happens when the file is manually resaved
            if item.data['LISTING ID'] == '' or item.data['LISTING ID'] == 'ID: ':
                item.data['LISTING ID'] = 'ID: ' + activeItem.data['Item ID']
                
            item.data['ebayTitle'] = activeItem.data['Item Title']
            item.data['STATUS'] = 'O'
            item.data['LISTING'] = 'LISTED'


        # Search in sold.csv for sold history. Note this search only applied to item not found in active.csv (hence the else)
        else:
            item.data['STATUS'] = 'OFFLINE'
            item.data['EBAY QTY'] = '0'
            flag, soldItem = Prod_Search(item, sold_List)

            # First assign the listing category ID if not already have one, whether listed or not
            if item.data['ebayCategory'] == '':
                item.data['ebayCategory'] = ebayCatNum[ item.data['PRODUCT'] ]
                
            # If the NON-ACTIVE items are found in sold listing, then do the following
            if flag == 'found':
                item.data['ebayID'] = 'ID: ' + soldItem.data['Item ID']
                #item.data['ebayID'] = 'Offline: ' + soldItem.data['Item ID']
                item.data['ebayTitle'] = soldItem.data['Item Title']
                #item.data['LISTING ID'] = 'Offline: ' + soldItem.data['Item ID']
                
                if item.data['ebayCategory'] == '':
                    item.data['ebayCategory'] = ebayCatNum[item.data['PRODUCT']]
                if item.data['LISTING ID'] == '' or item.data['LISTING ID'] == 'ID: ':
                    item.data['LISTING ID'] = 'ID: ' + soldItem.data['Item ID']
                    
                item.data['LISTING'] = 'RELIST'

            # For those not found in sold.csv nor active.csv but had a LISTING ID: relist (due to modified title). Else: UNLIST
            elif len(item.data['LISTING ID'].split()) == 2: #and item.data['LISTING ID'].split()[1] != '': # i.e. ID exists. not LISTING ID = '' , or, LISITING ID = 'ID: ' 
                item.data['LISTING'] = 'RELIST'
                
            else:
                item.data['LISTING'] = 'UNLIST'

        # Finally regardless of status, sold out is sold out
        if item.data['QTY'] == '0':
            item.data['LISTING'] = 'SOLD OUT'


    # Finally, record sold history
    for item in Prod_List:
        soldCount = 0

        for soldItem in sold_List:
            if len(item.data['LISTING ID'].split()) == 2 and item.data['LISTING ID'].split()[1] == soldItem.data['Item ID']:
                soldCount = soldCount + int(soldItem.data['Quantity'])
        item.data['SOLD'] = soldCount
    
    return Prod_List

def Prod_Search( item, ebay_List):

    flag = 'unfound'

    for ebayItem in ebay_List:
        #Two ways to search; by title or by ID
        if item.data['LISTING TITLE'].upper() == ebayItem.data['Item Title'].upper():
            flag = 'found'
            return flag, ebayItem
        else:
            if len(item.data['LISTING ID'].split()) == 2 and item.data['LISTING ID'].split()[1] == ebayItem.data['Item ID']:
                flag = 'found'
                return flag, ebayItem

    return flag, ebayItem
    
### Return flag, ebayItem
##def Prod_Search( item, ebay_List ):
##
##    flag = 'unfound'
##    searchKey = item.flexibleNames
##    #searchKey = item.IDs
##
##    # build use key
##    usekey = []
##    for name in searchKey:
##        if item.data[name] != '':
##            usekey = usekey + [item.data[name]+' '.upper()]
##            
##    #print usekey
##    size = len(usekey)
##
##    pos = dict()
##    neg = dict()
##    resultDict = dict()
##    
##    # Start searching against ebay_List
##    for ebayItem in ebay_List:
##        match = 0
##        false = 0
##        line = ebayItem.data['Item Title'].upper() #fix 'Item Title' in case it's not right
##
##
##        for key in usekey:
##
##            # positive find
##            if key in line:
##                match = match + 1
##
##
##            # No find
##            if key not in line:
##                false = false + 1
##
##
##        # Record result of curret ebayItem
##        pos[match] = ebayItem
##        neg[false] = ebayItem
##
##        
##    #print sorted(pos)
##
##
##    #print 'matchness: '+ str(max(pos)) + '/' + str(size) + ' = ' + str(float(max(pos))/float(size))
##    #print 'mismatch: '+ str(min(neg)) + '/' + str(size) + ' = ' + str(float(min(neg))/float(size)) 
##    
##
##
##    if neg[min(neg)].data['Item Title'] == pos[max(pos)].data['Item Title'] and min(neg) == 0:
##        return 'found', pos[max(pos)]
##    else:
##        return 'unfound', pos[max(pos)]





##def Prod_Search(item, ebay_List):
##    
##    # define a list of default search keys
##    searchKey = item.IDs
##    #searchKey = item.keyWords
##    #print searchKey
##    #searchKey = item.flexibleNames
##
##    count_dict = dict()
##    j = 0
##
##    # begin exploring each ebay item
##    for ebayItem in ebay_List:
##        j = j + 1
##        
##        matchCount = 0
##        falseCount = 0
##
##        # obtain the line to search in. the line is UPPER CASED
##        line = ebayItem.data['Item Title'].upper()
##
##        # keep the matchCount and recall the corresponding item.        
##        i = 0
##        while i < len(searchKey):
##
##            # Positive Count
##            if (item.data[ searchKey[i] ].upper() != '') and (' '+item.data[ searchKey[i] ].upper()+' ' in line): # use this line if searchKey is item.IDs or item.flexibleNames
##                matchCount = matchCount + 1
##
##            # Negative Count
##            if (item.data[ searchKey[i] ].upper() != '') and (item.data[ searchKey[i] ].upper() not in line):
##                falseCount = falseCount + 1
##            i = i + 1
##        #print j
##        if matchCount != 0:
##        #if matchCount != 0 and falseCount == 0:
##            count_dict[matchCount] = ebayItem
##            #print 'match count: ' + str(matchCount), ebayItem.data['Item Title'].upper()
##            
##    #print len(count_dict)
##    if len(count_dict) == 0:
##        return 'unfound', ebayItem
##    else:
##        #print max(count_dict), count_dict[max(count_dict)].data['Item Title']
##        return 'found', count_dict[max(count_dict)]


def load_data(filepath):
    
    # Read files and load data into memory (Assign excel data to PSU object)

    fin = open(filepath)
    ReadFile = csv.reader(fin)

    header = ReadFile.next() # header = ['MAKER', 'FIRM', ... ]
    PSU_inventory=[]
    
    for row in ReadFile:

        #initiate PSU class
        PSU_now = PSU()

        # Assign the header that this item has
        #PSU_now.giveHeader(header)
        PSU_now.dataNames = header

        # Assign the corresponding values of header
        i=0
        for item in row:
            PSU_now.dataValues = PSU_now.dataValues + [item.upper()]
            PSU_now.data[header[i]] = item.upper()
            i = i + 1
            
##        # Form key words for use in seaching against ebay files
##        words = []
##        for ID in PSU_now.IDs:
##            words = words + PSU_now.data[ID].split()
##            
##        i=0
##        while i < len(words):
##            # Remove the 0 in front of the ID values (useful for Dell's service tag that starts with 0; i.e. 0N58F)
##            if words[i][0] == '0':
##                temp = words[i][1:]
##            else:
##                temp = words[i]
##            s = ' ' + temp + ' ' # adding the spaces is a trick to modify key words to improve search robustness
##            PSU_now.keyWords = PSU_now.keyWords + [s]
##            i = i + 1

        # Collect into PSU_inventory

        PSU_inventory = PSU_inventory + [PSU_now]

        
    fin.close()

    return removeVoidEntries(PSU_inventory)

# Takes prud_list and the order of prints (by location or model).
# return arrays of excel extries such that new brands start at multiples of 47th lines. The separations are based on brands
def printerFormat_General(Sorted_Prod, heading):
    excelNewLine = 47
    collection = []
    brandNames = obtain_variation( Sorted_Prod, 'BRAND' )
    for brand in brandNames:

        vec47 = []
        vec47 = vec47 + [heading]

        #Build current brand
        brandArr = getSubList( Sorted_Prod, Sorted_Prod[0].data['PRODUCT'], brand ) # i.e. brandArrNow = getSubList( 'POWER SUPPLY', 'DELTA' )

        # keep building...
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

def removeVoidEntries(Prod_List):
    keep = []
    for item in Prod_List:
        i = 0
        true = 0

        # Traverse through the columns to see if there are non empty string
        while i < len(item.dataNames):
            if item.data[item.dataNames[i]] == '':
                true = true + 1
            i = i + 1


        if true != i: #all excel column is not all empty
            keep = keep + [item]
        
            
    return keep

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

def folder_check(all_paths, all_files):
    #all_paths = [backupPath, templatePath, offlinePath, printPath, inputPath, outputPath]
    #all_files = [active_ebayFile, sold_ebayFile, inventoryFile, templateFile]

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

def prodSort(Prod_List, instruction=0):

    # Determine the sorting order.
    if instruction == 0:
        sequence = ['PRODUCT', 'BRAND']
    else:
        sequence = instruction

    Prod_List = instructedSort(Prod_List, sequence)
    
    return Prod_List
    
def instructedSort(Prod_List, sequence): # Recursive sort. Note sequence = ['PRODUCT', 'BRAND', 'ID1']

    if len(sequence) == 1: # i.e. Prod_List = [item, ...etc]
        return sorted(Prod_List, key=lambda PSU_Object: PSU_Object.data[sequence[0]]) 
    else:
        result = []
        header = sequence[0] # i.e. header = 'PRODUCT'
        Prod_Dict = Dictionarize(Prod_List, header) # Prod_Dict = {'MOTHERBOARD': [item,...], 'POWER SUPPLY': [item,..]}

        #print sorted(Prod_Dict)

        for key in sorted(Prod_Dict): # i.e. key = 'MOTHERBOARD', or 'POWER SUPPLY' if header = 'PRODUCT'
            result = result + instructedSort(Prod_Dict[key], sequence[1:])
        return result
        

def Dictionarize(Prod_List, title): #title='PRODUCT' or 'BRAND' or 'NAMAE1' or 'ID1'...etc
    PartDict = dict()
    variationList = obtain_variation(Prod_List, title)

    for name in variationList:
        temp = []
        for item in Prod_List:
            if name == item.data[title]:
                temp = temp + [item]
        PartDict[name] = temp

    return PartDict

def obtain_variation(Prod_List, columnbTitle):
    i = 0
    variationArr = []
    #currentName = ''
    for item in Prod_List:
        if item.data[columnbTitle] not in variationArr:
            variationArr = variationArr + [item.data[columnbTitle]]
##        if item.data[columnbTitle] != currentName:
##            currentName = item.data[columnbTitle]
##            variationArr = variationArr + [currentName.upper()]
    variationArr.sort()
    return variationArr

class inventory:
    def __init__(self, myList):
        self.all = myList

    #def prodSort(self, sortOption='NONE'):
     #   if sortOption == 'NONE':
           # self.all = self.all


class PSU:
    def __init__(self):
        self.dataNames=[]
        self.dataValues=[]
        self.data = dict()
        self.flexibleNames = []
        self.hardNames = []
        self.IDs = []
        self.keyWords = []

    # Construct ebay listing title
    def eBayTitle(self):

        parts = []
        for title in self.flexibleNames:
            parts = parts + [self.data[title]]


        
        outStr = removeDSpace(' '.join(parts))
##
##        if len(outStr) > 80:
##            outStr = removeDSpace(' '.join(parts))
            
        i = 1
        while len(outStr) > 80:
            outStr = removeDSpace( ' '.join(parts[:-i]) )
            i = i + 1

        return outStr.upper()

    def warrantedTitle(self):
        s = self.eBayTitle()
        s1 = s+'Tested + Warranty'
        s2 = s+'Tested+Warranty'
        s3 = s+'+ Warranty'
        s4 = s+'+ Wrty'

        if len(s1) < 80:
            s = s1
        elif len(s2) < 80:
            s = s2
        elif len(s3) < 80:
            s = s3
        elif len(s4) < 80:
            s = s4
        else:
            pass
        
        return s

                       

##    def giveHeader(self, header):
##
##        # Assigns header to PSU object
##        self.dataNames = header
##
##
##        # If the inventory file is read (not active nor sold.csv) then there is a cut off column
##        cutoff = 'CUTOFF'
##        if cutoff in header:
##            Index = header.index('CUTOFF')
##            self.flexibleNames = header[: Index ]
##            self.hardNames = header[Index:]
##
##
##
##        # Collect IDs
##        ID = []
##        for name in header:
##            if 'ID' in name and name[len(name)-1].isdigit():
##                ID = ID + [name]
##        self.IDs = ID
##        
                
        
        
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
    
def getIndexedPath(filePath): # works with folder and file.
    
    pathNow = filePath
    i = 1
    if '.' in filePath:
        while os.path.exists(pathNow):
            tempLeft = filePath.split('.')[0]
            tempRight = filePath.split('.')[1]
            pathNow = tempLeft+'_'+str(i)+'.'+tempRight
            i = i + 1
        return pathNow
    #+'.'+filePath.split('.')[1]
    
    else:
        while os.path.exists(pathNow):
            pathNow = filePath+'_'+str(i)
            i = i + 1
        return pathNow

def makePath(folder):
    l = 1
    temp = folder
    while os.path.exists(folder):
        folder = temp + '_' + str(l)
        #print savePath
        l = l + 1
    os.makedirs(folder)
    

def apply_sortBy(Prod_List, file_inventory):
    
    print 'Sort MASTER INVENTORY by'
    print '0 - NONE'#. Enter NONE or just hit enter.'
    print '1 - ALPHA'#betical order of ebay title. Enter 1 or alpha.'
    print '2 - LOCATION'# order. Enter 2 or location '

    while True:
        
        answer = raw_input('Please enter one option above (Hit enter to default Model):').upper()

        # format answer
        if (answer == '0') or (answer == 'NONE') or (answer == ''):
            answer = 'NONE'
            
        if (answer == '1') or (answer == 'ALPHA'):
            answer = 'ALPHA'
            
        if (answer == '2') or (answer == 'LOCATION'):
            answer = 'LOCATION'

        # apply answer
        if answer == 'NONE' or answer == 'ALPHA' or answer == 'LOCATION':
            return listSort( Prod_List, answer )
                
        else:
            print 'Invalid input. Type the number (i.e. 1) or the name of sorting scheme (i.e. LOCATION).'
            pass

    #subprocess.Popen(r'explorer /select,'+ inventoryFile)





def menu():
    bar = '='*70
    sList = [
        bar,
        'Welcome to the General Inventory program',
        '',
        'Choose an option: ',
        '',
        '1 (print)   - Print Inventory',
        '2 (report)  - Generate offline reports and CSV listing files',
        '3 (sort)    - Choose how to sort Master Inventory',
        '4 (qty)     - Options to modify listing quantity',
        '',
        bar,
        '',
        ]
    i=0
    while i < len(sList):
        print sList[i]
        i = i + 1
    choice = raw_input('Choose option (example: enter 2 or report): ')
    return choice.upper()

    
def test():
    active_List = load_data(file_sold)
    keep = []
    output = []
    for item in active_List:
        if 'MOTHERBOARD' in item.data['Item Title'].upper() and item.data['Item Title'].upper() not in keep:
            keep = keep + [item.data['Item Title'].upper()]
            #output = output + [ [item.data['Item Title'].upper() , item.data['Item ID'], item.data['Price'], item.data['Quantity Available']]  ]
            output = output + [ [item.data['Item Title'].upper() , item.data['Item ID'], item.data['Sale Price']]  ]

    dst = 'C:\Users\JimmyQ6600\Desktop\here\\output.csv'
    CSV_ArrWrite( output, dst )
    print 'test done'
    sys.exit()

def test2(Prod_List):
    for item in Prod_List:
        print item.data['LISTING TITLE']
    sys.exit()

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
                print item.data['LOCATION'], ': ' + str(item.data['QTY']) + ' + ' + str(item.data['QTY UPDATE'])
                item.data['QTY'] = int(item.data['QTY']) + int(item.data['QTY UPDATE'])
                print item.data['QTY']
                item.data['QTY UPDATE'] = ''

    elif option == '4':
        file_recentlySold = 'c:\BLUESKY SCRIPTING\Input Files\\recentlySold.csv'
        recentlySold_List = load_data(file_recentlySold)
        for soldItem in recentlySold_List:

            # loop through our Prod_list inventory
            for myItem in Prod_List:
                if len(myItem.data['LISTING ID'].split()) == 2 and myItem.data['LISTING ID'].split()[1] == soldItem.data['Item ID']:
                    print myItem.data['LOCATION'], ': '+ str(myItem.data['QTY']) + ' - ' + str(soldItem.data['Quantity'])
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
        saveItemList(file_inventory, Prod_List, 10)
        print 'Modifications saved.'
        return Prod_List
    else:
        print 'Changes aborted. Returning to menu.'
        return Input_List


#test()
#test2()

####################
# Start of program
####################

folder_check(all_paths, all_files)
print 'Please wait while initializing...'

# Load files
Prod_List = load_data(file_inventory)
active_List = load_data(file_active)
sold_List = load_data(file_sold)

# Save a backup
file_backup = getIndexedPath( file_backup )
shutil.copy2( file_inventory, file_backup)
#saveBackup(srcDir, backupDir, fileName, fileName)
#saveItemList( file_backup, Prod_List)

# Update product online statuses and quantity
Prod_List = compare_ebay(Prod_List, active_List, sold_List)
Prod_List = listSort(Prod_List, 'LOCATION')

# Save updated
saveItemList(file_inventory, Prod_List, 0)


#subprocess.Popen(r'explorer /select,'+ file_backup)
print 'All entries read successfully. \n'

#test3(Prod_List, active_List, sold_List)

choice = ""
while True:
#Note, all options must take care to save to inventory on its own if they happen to temper with the Prod_List

    # Get user input for choice
    choice = menu()
    if choice == '1' or choice == 'PRINT':
        print_inventory(Prod_List, printPath)
        #dummy1 = raw_input('Hit enter to return to menu')
        
    elif choice == '2' or choice == 'REPORT':
        print 'Please wait while report files are being generated... '
        csv_report(Prod_List, active_List, sold_List, path_offline, templatePath )
        #dummy2 = raw_input('Hit enter to return to menu')
        
    elif choice == '3' or choice == 'SORT':
        Prod_List = apply_sortBy(Prod_List, file_inventory)
        saveItemList(file_inventory, Prod_List, 10)
        subprocess.Popen(r'explorer /select,'+ file_inventory)
        #set_sortBy_flag()
        #dummy3 = raw_input('Hit enter to return to menu')
        
    elif choice == '4' or choice == 'QTY':
        Prod_List = modQTY(Prod_List, active_List, sold_List)
        #dummy4 = raw_input('Hit enter to return to menu')
        
    elif choice == '5' or choice == 'mod':
        modify_entry()
        #dummy4 = raw_input('Hit enter to return to menu')
        
    elif choice == '6' or choice == 'exp':
        export_entry()
        #dummy4 = raw_input('Hit enter to return to menu')

    elif choice == '7' or choice == 'EXIT':
        break

    print ''

