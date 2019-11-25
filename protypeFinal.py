# Author: Wahid Rahimi
# Information about the array produced from this program:

# ____________________________________________________________________________________________________________________
#     # itemsArrayFormat shows the format of how EACH item array will look
#     #itemArrayFormat = ["Item#", "Quantity", "8A Quantity", "SA Quantity", "Unrestricted", "Period of Performance",
#     #                   "NSN", "Delivery Identification", "State", "Region", "Throughput SPLC", "Requirement SPLC",
#     #                  "Delivery Address",
#     #                    "Service Code", "Delivery DODAAC", "Ordering Office DODAAC", "Notes"]
#
#     # modesFormat is another array because there can be multiple modes in each item
#    # modesFormat = ["Mode", "Reciept %", "Max Parcel", "MinParcel", "FOB Restriction", "FSII", "SDA", "CI"]
#
#     # Delivery Notes for each item will be one single string
#     #sometimes = "Delivery Notes"
# ____________________________________________________________________________________________________________________


import PyPDF2
import re

def main():
    filename = 'WESTPAC_sol.pdf'
    clean_result = dataExtraction(filename)
    mainlist = []
    for i in range(len(clean_result)):
        item = clean_result[i]
        #print(f"Item {i}: ")
        #print(item.split())

        itemattributes = []
        #print("0 to nsn")
        #print(extractToNSN(item))
        itemattributes += ((extractToNSN(item)))
        print()
        #print("nsn to region")
        #print(extractNSNtoRegion(item))
        itemattributes += (extractNSNtoRegion(item))
        print()
        #print("Region to Service Code:")
        #print(extractRegionToService(item))
        itemattributes += (extractRegionToService(item))
        print()
        #print("Service Code to Mode")
        #print(extractServiceCodeToMode(item))
        itemattributes += (extractServiceCodeToMode(item))
        #print()
        mainlist.append(itemattributes)



    print("THIS IS MAINLIST")
    print(mainlist[0][6:10])
    print(mainlist[1][6:10])
    print(mainlist[2][6:10])
    print(mainlist[12][6:10])


def dataExtraction(filename):
    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    numPages = pdfReader.numPages - 62

    everyPage = ''
    for pageNum in range(8, numPages):
        # Create a temporary page object
        tempPage = pdfReader.getPage(pageNum)
        # Extracts text from current page and adds it to everyPage
        everyPage += tempPage.extractText()

    indexOfFirst = everyPage.find("Item: 00")
    onlyItems = everyPage[indexOfFirst:]
    result = onlyItems.split("Item: ")
    result.pop(0)

    # clean_result returns an array without the "T" items.
    clean_result = []
    index = 0
    for index in range(len(result)):
        if (result[index][:2] == ' T'):
            continue
        else:
            clean_result.append(result[index])

    return clean_result

def extractToNSN(item):
    item0tonsn = item[0:item.find("NSN")]
    array0tonsn = item0tonsn.split()
    result = []
    # Item
    result.append(array0tonsn[0])
    # Quantity
    result.append(array0tonsn[2] + " " + array0tonsn[3])
    # 8A Quantity
    result.append(array0tonsn[6])
    # SA Quantity
    result.append(array0tonsn[9])
    # Unrestricted
    result.append(array0tonsn[11])
    # Period of Performence
    result.append(array0tonsn[15])

    return result

def extractNSNtoRegion(item):
    focus = item[item.find("NSN"):item.find("Region")]
    result = ["NSN", "DeliveryID", "State"]

    if (focus[focus.find("State") + 6] == "\n"):
        flag = True
        start = focus.find("State")
        while (flag):
            num_nl = focus.count("\n")
            if (num_nl > 1):
                focus = focus[num_nl + 1:]
            else:
                flag = False
    else:
        focus = focus[focus.find("State") + 6:]
    result[0] = focus[:focus.find(")") + 1]
    result[1] = focus[focus.find(")") + 1:focus.find("\n")]
    thirdline = focus[focus.find("\n") + 1:]
    if (thirdline[2] == " "):
        result[2] = thirdline[:2]
        result[0] += " " + thirdline[2:]
    else:
        result[2] = ""
        result[0] += " " + thirdline
    return result

def extractRegionToService(item):
    result = ["Region", "ThroughputSPLC", "RequirementSPLC", "Delivery Address"]
    focus = item[item.find("Region"):item.find("Service Code")]

    splitfoc = focus.split()
    if (focus.find("Throughput") > -1):

        result[0] = splitfoc[4][:2]
        result[1] = splitfoc[4][2:]
        result[2] = splitfoc[5]
        result[3] = focus[focus.find("Address:") + 9:]
        return result
    else:

        result[0] = splitfoc[2][:2]
        result[1] = ""
        result[2] = splitfoc[2][2:]
        result[3] = focus[focus.find("Address:") + 9:]
        return result

def extractServiceCodeToMode(item):
    focus = item[item.find("Service Code"):item.find("Mode")]

    result = ["Service Code", "Delivery DODAAC", "Ordering Office DODAAC"]
    focus_split = focus.split()
#    Checks if there is a Service Code
    if (focus.find("SE") == 0):
        result[0] = ""
    else:
        result[0] = focus[focus.find("\n") + 1:focus.find("SE")]
    result[1] = focus[focus.find("SE"):focus.find("SE") + 6]
    lastElem = focus_split[len(focus_split) - 1]
    lenDODAAC = len(lastElem[lastElem.find("SE"):])
#    Checks to see if there is an Ordering Office DODAAC
    if (lenDODAAC > 6):
        result[2] = focus[focus.find("SE") + 6:]
    else:
        result[2] = ""
    return result



if __name__ == '__main__':
    main()
