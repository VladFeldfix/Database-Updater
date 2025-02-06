# this file gather information for the Products spreadsheet
# the fields displayed will be the following:

# 1. Product Part Number    for example BB1K200C
# 2. Plug Name              for example J1
# 3. Plug Part Number       for example D38999/24D35SN
# 4. Connect to Test Cable  for example R1_189_1

import os

class Products:
    def get_data(programs):
        result = [("Product Part Number","Plug Name","Plug Part Number","Connect to Test Cable")]
        products = {} # BB1K200C: [ (J1, D38999/24D35SN, R1_189_1) ]
        for root, dirs, files in os.walk(programs):
            # get product part number
            product_part_number = root.replace(programs, "")
            product_part_number = product_part_number.replace("\\", "/")
            product_part_number = product_part_number.split("/")
            if len(product_part_number) > 1:
                product_part_number = product_part_number[1]
            
            # get testcables to product files
            for file in files:
                if "testcables_to_product.csv" in file:
                    f = open(root+"/"+file, 'r')
                    lines = f.readlines()
                    f.close()
                    for line in lines[1:]:
                        line = line.replace("\n", "")
                        line = line.split(",")
                        if len(line) > 2:
                            if not product_part_number in products:
                                products[product_part_number] = []
                            products[product_part_number].append((line[1],line[2],line[0]))
        
        # generate result
        for product_part_number, details in products.items():
            for detail in details:
                # ("Product Part Number","Plug Name","Plug Part Number","Connect to Test Cable")
                plugname = detail[0]
                plugpn = detail[1]
                connectto = detail[2]
                if detail[0] == "":
                    plugname = "Missing"
                if detail[1] == "":
                    plugpn = "Missing"
                if detail[2] == "":
                    connectto = "Missing"
                result.append((product_part_number, plugname, plugpn, connectto))
        return result