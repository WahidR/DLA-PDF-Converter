# Author: Wahid Rahimi
# Information about the array produced from this program:

# ____________________________________________________________________________________________________________________
#     # itemsArrayFormat shows the format of how EACH item array will look
#     #itemArrayFormat = ["Item#", "Quantity", "8A Quantity", "SA Quantity", "Unrestricted", "Period of Performance",
#     #                   "NSN", "Delivery Identification", "State", "Region", "Throughput SPLC", "Requirement SPLC",
#     #                  "Delivery Address",
#     #                   "Service Code", "Delivery DODAAC", "Ordering Office DODAAC", "[mode1, mode2, notes]"]
#
#
#     # modesFormat is another array because there can be multiple modes in each item. Each mode will be in a mode list
#    # modesFormat = ["Mode", "Reciept %", "Max Parcel", "MinParcel", "FOB Restriction", "FSII", "SDA", "CI"]
#
#    # mode list will be [mode1, mode2, notes] or [mode1, notes]
#     Delivery Notes and Delivery Hours for each item will be one single string and will be in the modeinfo list
#     #notes = "Delivery Hours: 7AM - 12PM Delivery Notes: blablabla"
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

        #print("nsn to region")
        #print(extractNSNtoRegion(item))
        itemattributes += (extractNSNtoRegion(item))

        #print("Region to Service Code:")
        #print(extractRegionToService(item))
        itemattributes += (extractRegionToService(item))

        #print("Service Code to Mode")
        #print(extractServiceCodeToMode(item))
        itemattributes += (extractServiceCodeToMode(item))

        #print("Mode to End")
        #print(extractModeToEnd(item))
        #print("\n")
        itemattributes += (extractModeToEnd(item))

        mainlist.append(itemattributes)



    print("THIS IS MAINLIST")
    print(mainlist[0])
    print(mainlist[1])



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

def maxandminfinder(numbers):
    result = ["max", "min"]
    firstnum = ""
    secondnum = ""
    orginialnum = numbers
    index = numbers.find(",")
    foundfirstnum = False
    if (numbers.find(",") < 0 and len(numbers) <= 3):
        firstnum = numbers
        secondnum = ""
    elif (numbers.find(",") < 0):
        firstnum = numbers[:3]
        secondnum = numbers[3:]
    else:

        while (True):

            commalocation = numbers.find(",")

            if (numbers[index:].find(",") < 0):
                firstnum = orginialnum
                break
            elif (numbers.find(",") > 3):
                endindexoffirstnum = len(orginialnum) - len(numbers) + 3
                firstnum = orginialnum[:endindexoffirstnum]
                secondnum = orginialnum[endindexoffirstnum:]
                break
            else:
                numbers = numbers[numbers.find(",") + 1:]
    result[0] = firstnum
    result[1] = secondnum

    return result


def extractModeToEnd(item):
    focus = item[item.find("BULK:"):]


    #   listofmodes = [mode1,mode2,delivery instructions]
    #   len(listofmodes) will always be the number of modes + 1 b/c of delivery instructions
    listofmodes = []

    foclist = focus.split("BULK:")[1:]
    # This loops through each mode
    for i in range(len(foclist)):
        currentmode = foclist[i]
        splitmode = currentmode.split()
        modeinfo = ["Mode", "Reciept%", "Max Parcel", "Min Parcel", "FOB Restriction", "FSII", "SDA", "CI"]

        # This will repeat the extracting process according to the amount of modes there are
        for j in range(len(splitmode)):
            # Run everything under this
            if len(foclist[i].split()) == 0:
                continue
            # This will always end up as
            elif (len(foclist[i].split()) == 1):
                modeinfo[0] = currentmode[:currentmode.find("100")]
                # Reciept% is always 100
                modeinfo[1] = "100"

                if (currentmode[currentmode.find("100") + 3].isalpha()):
                    modeinfo[2] = "0"
                    modeinfo[3] = "0"
                    modeinfo[4] = currentmode[currentmode.find("100") + 3]
                    if (currentmode[len(currentmode) - 1] == 'Y'):
                        modeinfo[5] = "Y"
                    elif (currentmode[len(currentmode) - 1] == 'N'):
                        modeinfo[5] = "N"
                    else:
                        modeinfo[5] = ""

                else:
                    reverse_modeinfo1 = splitmode[0][len(splitmode[0]) - 1:0]
                    endindex = len(splitmode[0]) - (reverse_modeinfo1.find("0"))  # Shouldn't need a plus one
                    numbers = currentmode[currentmode.find(100) + 3:endindex]
                    min_max_array = maxandminfinder(numbers)
                    modeinfo[2] = min_max_array[0]
                    modeinfo[3] = min_max_array[1]
                    modeinfo[4] = currentmode[endindex:endindex + 1]  # It is only the end index for the numbers slice
                    modeinfo[5] = ""
                modeinfo[6] = ""
                modeinfo[7] = ""



            else:
                # modeinfo = ["Mode", "Reciept%", "Max Parcel", "Min Parcel", "FOB Restriction", "FSII", "SDA", "CI"]

                modeinfo[0] = splitmode[0][:splitmode[0].find("100")]
                # Reciept% is always 100
                modeinfo[1] = "100"
                remainingstring = splitmode[0][splitmode[0].find("100")]
                reverse_modeinfo1 = splitmode[0][::-1]
                endindex = len(splitmode[0]) - (reverse_modeinfo1.find("0"))  # Shouldn't need a plus one

                numbers = currentmode[currentmode.find("100") + 3:endindex]
                min_max_array = maxandminfinder(numbers)

                modeinfo[2] = min_max_array[0]
                modeinfo[3] = min_max_array[1]
                modeinfo[4] = currentmode[endindex:endindex + 1]
                modeinfo[5] = splitmode[0][len(splitmode[0]) - 1:]
                modeinfo[6] = splitmode[1]
                modeinfo[7] = splitmode[2]

        listofmodes.append(modeinfo)

    # Delivery Notes extraction:

    if (focus.find("Delivery") >= 0):
        delivery_notes = focus[focus.find("Delivery"):]
    else:
        delivery_notes = "No Delivery Notes"
    listofmodes.append(delivery_notes)

    return listofmodes


if __name__ == '__main__':
    main()
