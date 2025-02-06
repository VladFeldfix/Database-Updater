# this file gather information for the Test Logs spreadsheet
# the fields displayed will be the following:

# 1. Product Part Number    for example BB1K200C
# 2. Serial Number          for exaamle FLT2425-0748
# 3. Date                   for example 20Jan2025

import os

class Test_Logs:
    def get_data(path_test_results):
        result = [("Product Part Number","Serial Number","Date")]

        for root, dirs, files in os.walk(path_test_results):
            for file in files:
                file = file.split(".")
                file = file[0]
                file = file.split("_")
                if len(file) == 3:
                    result.append((file[2],file[0],file[1]))
        return result