#-------------------------------------------------------------------------------
# Name:        main
# Revision:    7.0
#
# Purpose:     Developed for Bluesky Trading.
#              Inventory control, inventory look up and updates,
#               parse eBay files for product online statuses, output eBay instruction files.
#
# Author:      Chi Yung 'Jimmy' Fung
#
# Created:     24/12/2013
# Copyright:   (c) Bluesky Trading 2013
# Licence:     <Open Source>
#-------------------------------------------------------------------------------

# Note for importing global variables (file paths, configs...etc)
# Python does not truly allow global variables in the sense that C/C++ do.
# An imported Python module will not have direct access to the globals in the module which imports it, nor vice versa.


# Imrpovement ideas

# - Programatically upload ebayfileExchange (see java sample)
# - host web server (apache, LAMP), host own images.
# - Repond to ebay upload result file, unfound listing should be treated as unlist
# - Display the search key used to locate the entry
# - Integrity check: relates to ID and sales
# - Scanner mode


#View ebay categor
#http://listings.ebay.com/_W0QQloctZShowCatIdsQQsocmdZListingCategoryList


ebayCatNum = dict()
ebayCatNum['']                = ''
ebayCatNum['POWER SUPPLY']    = 42017
ebayCatNum['MOTHERBOARD']     = 1244
ebayCatNum['VIDEO CARD']      = 175673
ebayCatNum['CAR PART']        = 6001
ebayCatNum['CPU']             = 164
ebayCatNum['HARD DRIVE']      = 165


dcpt_html = ['Tested working.<div>Guaranteed against DOA.</div><div>30 days warranty.</div><div><br></div><div>Continental US buyers only.</div><div>Please check my listing for more computer parts.</div>',
             'Item sold as is.<div><br><div>Continental US buyers only.</div><div>Please check my listing for more computer parts.</div></div>',
             ]
s_header = 'Action(CC=Cp1252),SiteID,Format,Title,Condition,ConditionDescription,SubTitle,Custom Label,Category,Category2,StoreCategory,StoreCategory2,Quantity,LotSize,Currency,StartPrice,BuyItNowPrice,ReservePrice,InsuranceOption,InsuranceFee,DomesticInsuranceOption,DomesticInsuranceFee,PackagingHandlingCosts,InternationalPackagingHandlingCosts,Duration,PrivateAuction,Country,ProductIDType,ProductIDValue,Product:ProductReferenceID,ItemID,Description,HitCounter,PicURL,BoldTitle,Featured,GalleryType,FeaturedFirstDuration,Highlight,Border,HomePageFeatured,Subtitle in search resutls,GiftIcon,GiftServices-1,GiftServices-2,GiftServices-3,SalesTaxPercent,SalesTaxState,ShippingInTax,UseTaxTable,PostalCode,ProxyItem,VATPercent,Location,ImmediatePayRequired,PayPalAccepted,PayPalEmailAddress,PaymentInstructions,PaymateAccepted,ProPayAccepted,MoneyBookersAccepted,StandardPayment,UPC,CCAccepted,AmEx,Discover,VisaMastercard,IntegratedMerchantCreditCard,COD,CODPrePayDelivery,PostalTransfer,MOCashiers,PersonalCheck,MoneyXferAccepted,MoneyXferAcceptedinCheckout,PaymentOther,OtherOnlinePayments,PaymentSeeDescription,Escrow,ShippingType,GlobalShippingService,ShipFromZipCode,ShippingIrregular,ShippingPackage,WeightMajor,WeightMinor,WeightUnit,MeasurementUnit,ShippingDetails/CODCost,PackageLength,PackageWidth,PackageDepth,DomesticRateTable,InternationalRateTable,CharityID,CharityName,DonationPercent,ShippingService-1:Option,ShippingService-1:Cost,ShippingService-1:AdditionalCost,ShippingService-1:Priority,ShippingService-1:FreeShipping,ShippingService-1:ShippingSurcharge,ShippingService-2:Option,ShippingService-2:Cost,ShippingService-2:AdditionalCost,ShippingService-2:Priority,ShippingService-2:ShippingSurcharge,ShippingService-3:Option,ShippingService-3:Cost,ShippingService-3:AdditionalCost,ShippingService-3:Priority,ShippingService-3:ShippingSurcharge,ShippingService-4:Option,ShippingService-4:Cost,ShippingService-4:AdditionalCost,ShippingService-4:Priority,ShippingService-4:ShippingSurcharge,ShippingService-5:Option,ShippingService-5:Cost,ShippingService-5:AdditionalCost,ShippingService-5:Priority,ShippingService-5:ShippingSurcharge,GetItFast,DispatchTimeMax,IntlShippingService-1:Option,IntlShippingService-1:Cost,IntlShippingService-1:AdditionalCost,IntlShippingService-1:Locations,IntlShippingService-1:Priority,IntlShippingService-2:Option,IntlShippingService-2:Cost,IntlShippingService-2:AdditionalCost,IntlShippingService-2:Locations,IntlShippingService-2:Priority,IntlShippingService-3:Option,IntlShippingService-3:Cost,IntlShippingService-3:AdditionalCost,IntlShippingService-3:Locations,IntlShippingService-3:Priority,IntlShippingService-4:Option,IntlShippingService-4:Cost,IntlShippingService-4:AdditionalCost,IntlShippingService-4:Locations,IntlShippingService-4:Priority,IntlShippingService-5:Option,IntlShippingService-5:Cost,IntlShippingService-5:AdditionalCost,IntlShippingService-5:Locations,IntlShippingService-5:Priority,IntlAddnlShiptoLocations,PaisaPayAccepted,PaisaPay EMI payment,BasicUpgradePackBundle,ValuePackBundle,ProPackPlusBundle,BestOfferEnabled,AutoAccept,BestOfferAutoAcceptPrice,AutoDecline,MinimumBestOfferPrice,BestOfferRejectMessage,LocalOnlyChk,LocalListingDistance,BuyerRequirements:ShipToRegCountry,BuyerRequirements:ZeroFeedbackScore,BuyerRequirements:MinFeedbackScore,BuyerRequirements:MaxUnpaidItemsCount,BuyerRequirements:MaxUnpaidItemsPeriod,BuyerRequirements:MaxItemCount,BuyerRequirements:MaxItemMinFeedback,BuyerRequirements:LinkedPayPalAccount,BuyerRequirements:VerifiedUser,BuyerRequirements:VerifiedUserScore,BuyerRequirements:MaxViolationCount,BuyerRequirements:MaxViolationPeriod,SellerDetails:PrimaryPhone,SellerDetails:SecondaryPhone,ExtSellerDetails:Hours1Days,ExtSellerDetails:Hours1AnyTime,ExtSellerDetails:Hours1From,ExtSellerDetails:Hours1To,ExtSellerDetails:Hours2Days,ExtSellerDetails:Hours2AnyTime,ExtSellerDetails:Hours2From,ExtSellerDetails:Hours2To,ExtSellerDetails:TimeZoneID,ListingDesigner:LayoutID,ListingDesigner:ThemeID,ShippingDiscountProfileID,InternationalShippingDiscountProfileID,Apply Profile Domestic,Apply Profile International,PromoteCBT,ShipToLocations,CustomLabel,CashOnPickup,ReturnsAcceptedOption,ReturnsWithinOption,RefundOption,ShippingCostPaidBy,ReturnsRestockingFee,WarrantyOffered,WarrantyType,WarrantyDuration,AdditionalDetails,MarketplaceType,ProjectGoodCategory,ShortDescription,ProducerDescription,RegionOfOrigin,ProducerPhotoURL,Relationship,RelationshipDetails'
s_values = 'Add,US,FixedPriceItem,example,3000,,,,42017,,1,0,1,,USD,57.99,,,,,,,,,GTC,0,US,,,,,@@@@%<P>Pulled from working systems.</P>%0D%0A<P>Tested working.</P>%0D%0A<P>Comes with 30 days warranrty.</P>%0D%0A<P>Guarantee against DOA</P>%0D%0A<P>Continental US buyers only.</P>%0D%0A<P>Please see my listing for more computer parts.</P>,,,0,0,None,,0,0,0,0,0,,,,0.00,,,0,27603,,,,1,1,wei2cool4u2004@gmail.com,None Specified,,,,,,,,,,,,,,,,,,,,0,,Flat,0,,,Letter,5,,,English,,,,,,,,,,ShippingMethodStandard,0.00,,1,1,,,,,,,,,,,,,,,,,,,,,,0,1,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,0,,,,,,,,,,1,,-1,2,Days_30,,,1,,,4,Days_30,,,,,,,,,,,,,,0||,0||,0,0,,,,,ReturnsAccepted,Days_30,MoneyBack,Buyer,NoRestockingFee,,,,,,,,,,,,'
template_header = s_header.split(',')
template_values = s_values.split(',')


# Essential Folders
root = 'c:\BLUESKY SCRIPTING'
backupPath = root + '\\Backup'
templatePath = root + '\\Listing Template'
offlinePath = root + '\\Offline Listing'
printPath = root + '\\Printed Inventory'
inputPath = root + '\\Input Files'
outputPath = root + '\\Output Files'
all_paths = [backupPath, templatePath, offlinePath, printPath, inputPath, outputPath]

# Input Files
file_active = root + '\\Input Files\\active.csv'
file_sold = root + '\\Input Files\\sold.csv'
file_inventory = root + '\\Input Files\\Master Template.csv'
#file_template = root + '\\Listing Template\\PSU.csv'
all_files = [file_active, file_sold, file_inventory] 

# Output paths
path_offline = root + '\\Offline Listing\Offline Output'

# Output Files
file_backup = root + '\\Backup\\Master Template.csv'

paths = {
    'root'           : root,
    'backupPath'     : backupPath,
    'templatePath'   : templatePath,
    'offlinePath'    : offlinePath,
    'printPath'      : printPath,
    'inputPath'      : inputPath,
    'outputPath'     : outputPath,
    'file_active'    : file_active,
    'file_sold'      : file_sold,
    'file_inventory' : file_inventory,
    'path_offline'   : path_offline,
    'file_backup'    : file_backup,
    'ebayCatNum'     : ebayCatNum,
    'template_header': template_header,
    'template_values': template_values,
}      


from module import *


if __name__ == '__main__':

    folder_check(all_paths, all_files)

    print 'Please wait while loading inventory...'

    # Load files and save a backup of loaded file
    Prod_List, active_List, sold_List = initialize(paths)

    # Update product online statuses and quantity
    Prod_List = compare_ebay(Prod_List, active_List, sold_List)
    Prod_List = listSort(Prod_List, 'location') #buggy
    
    saveItemList(file_inventory, Prod_List, 0)                  # Save updated
    print 'All entries read successfully. \n'


    # Startup1
    startup1(Prod_List, active_List, sold_List)
