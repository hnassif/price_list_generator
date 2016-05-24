# -*- coding: utf-8 -*- 
import sys
import codecs
import pandas as pd

sys.stdout = codecs.getwriter('utf8')(sys.stdout)


def get_cis_products_list(filename,sheetname):

    # Read and Parse the 'Remove_Dup' sheet in the Mapping Produts Excel File
    xl = pd.ExcelFile(filename)
    df = xl.parse(sheetname)

    # Extract the 'HP_Reference_ListP' from the file
    hp_reference_list = df['HP_Reference_ListP']

    # Remove empty rows (This is to know where the data ends)
    criterion = df[ pd.notnull(df['HP_Reference_ListP']) ] 

    # Display the Size of the data 
    print( 'CIS Product List Dataframe Size : ' + str(criterion.size) )
    print( 'CIS Product List Dataframe Shape : ' + str(criterion.shape) )

    #The number of rows corresponds to the number of unique products 
    num_of_rows = criterion.shape[0] 
    print( 'Number of Products : ' + str(num_of_rows-1)) # -1 because of the first row, which has the column names

    # Index the DataFrame. ie: assign numbers to rows 
    df.set_index([range(df.shape[0])], inplace = True)

    hp_products_purchased = set(criterion['HP_Reference_ListP'])

    return hp_products_purchased


# Function to remove Non-ASCII characters. This function is used in the 'generate_filtered_product_list' function.
def removeNonAscii(s): 
	return "".join(map(lambda x: x if ord(x)<128 else 'e', s))

 
def generate_filtered_product_list(hp_price_list,sheetname,hp_products_purchased):

    # Read and Parse the 'sheet1' sheet in the Mapping Produts Excel File
    xl_pl = pd.ExcelFile(hp_price_list)
    df_pl = xl_pl.parse(sheetname)

    # Converting columns to ASCII characters, because HP uses some non-ascii characters that the Panda library cannot handle
    columns = df_pl.columns
    df_pl.columns = map( removeNonAscii , list(columns) )

    # Remove empty rows (This is to know where the data ends)
    criterion_pl = df_pl[ pd.notnull(df_pl['Numero de produit']) ] 

    # Display the Size of the data 
    print( 'HP Product List Size : ' + str(criterion_pl.size) )
    print( 'HP Product List Shape : ' + str(criterion_pl.shape) )

    #The number of rows corresponds to the number of HP Prodcuts
    num_of_rows_pl = criterion_pl.shape[0]
    print('Number of Rows in HP Price List : ' + str(num_of_rows_pl) )

    # 'frames' will store the results of each product query performed
    frames = []

    # For each hp_product purchased by CIS, search for any HP product that starts with the same number 
    for hp_product in hp_products_purchased:
        query = df_pl['Numero de produit'].map(lambda x: x.startswith(hp_product) if type(x) != float else False)
        if df_pl[query].empty == False :
            frames += [ df_pl[query] ]

    # Combining all query resuts into a single DataFrame  
    output_df = pd.concat(frames)
    print( 'The filtered HP products list has Size : ' + str(output_df.size) )
    print( 'The filtered HP products list has Shape : ' + str(output_df.shape) )


    #Writing the Output to an excel file 
    writer = pd.ExcelWriter('output_' + hp_price_list + '.xlsx')
    output_df.to_excel(writer,'Sheet1', index=False)
    writer.save()

    return 










