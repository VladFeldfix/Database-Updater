# this file gather information for the Programs spreadsheet
# the fields displayed will be the following:

# 1. Product Part Number    for example BB1K200C
# 2. MPT5000L               Complete / In progress
# 3. MPT5000L Validation    Y/N
# 4. MPT5000                Complete / In progress
# 5. MPT5000 Validation     Y/N

import os

class Programs:
    def get_data(programs, results):
        result = [("Product Part Number","MPT5000L","MPT5000L Validation","MPT5000","MPT5000 Validation")]
        products = {} # product_part_number: ["MPT5000L","MPT5000L Validation","MPT5000","MPT5000 Validation"]
        
        # go to the programs folder to get information about what programs are in progress
        for root, dirs, files in os.walk(programs):
            product_part_number = root.replace(programs, "")
            product_part_number = product_part_number.replace("\\", "/")
            product_part_number = product_part_number.split("/")
            if len(product_part_number) > 1:
                product_part_number = product_part_number[1]
                if not product_part_number in products:
                    products[product_part_number] = ["In Progress", "N", "In Progress", "N"]

        # go to the the test results folder and see what programs are completed
        for root, dirs, files in os.walk(results):
            product_part_number = root.replace(results, "")
            product_part_number = product_part_number.replace("\\", "/")
            product_part_number = product_part_number.split("/")
            if len(product_part_number) > 1:
                product_part_number = product_part_number[1]
            
                for file in files:
                    # testing for MPT5000 program and validation file
                    if ".mpt_product" in file:
                        if product_part_number in products:
                            products[product_part_number][2] = "Complete"
                            if os.path.isfile(results+"/"+product_part_number+"/Test Program/MPT5000/"+product_part_number+" MPT5000-Validation.pdf"):
                                products[product_part_number][3] = "Y"
                    
                    # testing for MPT5000 program and validation file
                    if product_part_number+".txt" in file:
                        if product_part_number in products:
                            products[product_part_number][0] = "Complete"
                            if os.path.isfile(results+"/"+product_part_number+"/Test Program/MPT5000L/"+product_part_number+" MPT5000L-Validation.pdf"):
                                products[product_part_number][1] = "Y"
        
        # add gathered data to result
        for key, val in products.items():
            result.append((key,val[0],val[1],val[2],val[3]))
        
        # return results
        return result