import os 
import price_list_generator as plg

# 'all_price_lists' will store the name of all the PL files in the current directory
all_price_lists = [] 

# iterate over all the PL files in the current directory and add their names to 'all_price_lists'
for i in os.listdir(os.getcwd()):
    if i.endswith(".xlsx") and i.startswith('PL'): 
        all_price_lists += [i]
        continue
    else:
        continue

# Get all products purchased
hp_products_purchased = plg.get_cis_products_list("Mapping_Products_2.xlsx","Remove_Dup")


# Filter every HP price list in 'all_price_lists'
for pl in all_price_lists:
	print('Working on PL : ' + pl + ' ...')
	plg.generate_filtered_product_list(pl,"SHEET1",hp_products_purchased)
	print('Done with PL : ' + pl)



