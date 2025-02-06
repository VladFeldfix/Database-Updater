# this file gather information for the Jigs spreadsheet
# the fields displayed will be the following:

# 1. Jig Id         for example J12
# 2. Test Cables    for example: R1_047, R1_048 (each one a separate row)

class Jigs:
    def get_data(jigs):
        result = [("Jig Id","Test Cables")]
        file = open(jigs, 'r')
        lines = file.readlines()
        file.close()
        for line in lines:
            line = line.split(",")
            if len(line) == 2:
                result.append((line[0], line[1].replace("\n", "")))
        return result
    