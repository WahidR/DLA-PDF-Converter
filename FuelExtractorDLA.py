import PyPDF2
import xlsxwriter


class Mode:

    def __init__(self, modeType="", reciept="", maxParcel="", minParcel="",
                 fobRestriction="", fsii="", sda="", ci=""):
        if (modeType == ""):
            modeType = "N/A"
        if (reciept == ""):
            reciept = "N/A"
        if (maxParcel == ""):
            maxParcel = "N/A"
        if (minParcel == ""):
            minParcel = "N/A"
        if (fobRestriction == ""):
            fobRestriction = "N/A"
        if (fsii == ""):
            fsii = "N/A"
        if (sda == ""):
            sda = "N/A"
        if (ci == ""):
            ci = "N/A"

        self.modeType = modeType
        self.receipt = reciept
        self.maxParcel = maxParcel
        self.minParcel = minParcel
        self.fobRestriction = fobRestriction
        self.fsii = fsii
        self.sda = sda
        self.ci = ci

    def __repr__(self):
        output = "Mode Type : " + self.modeType + " Reciept%: " + self.receipt + " Max Parcel: " + self.maxParcel
        output += " Min Parcel: " + self.minParcel + " FOB Restriction: " + self.fobRestriction
        output += " FSII: " + self.fsii + " SDA: " + self.sda + " CI: " + self.ci

        return output

    def __str__(self):
        output = "" + self.modeType + " " + self.receipt + " " + self.maxParcel
        output += " " + self.minParcel + " " + self.fobRestriction
        output += " " + self.fsii + " " + self.sda + " " + self.ci

        return output


class ModeList:

    def __init__(self, modes=None):
        if modes is None:
            modes = []
        if isinstance(modes, Mode):
            modes = [modes]
        self.modes = modes

    def add(self, mode):
        if isinstance(mode, Mode):
            self.modes += [mode]
        else:
            self.modes += mode

    def remove(self, mode):
        self.remove(mode)

    def pop(self, index=None):
        if index is None:
            return self.modes.pop()

        return self.modes.pop(index)

    def __str__(self):
        output = "["
        for i in range(len(self.modes)):
            output += str(self.modes[i]) + ", "
        output = output[:len(output) - 2]
        output += "]"
        return output

    def __repr__(self):
        output = ""
        for i in range(len(self.modes)):
            output += str(self.modes[i]) + ", "
        output = output[:len(output) - 2]
        return output


class Item:
    def __init__(self, itemNum, quantity, quantity8A, quantitySA, unrestricted, performancePeriod, nsn, deliveryID,
                 state, region, throughputSPLC, requirementSPLC, deliveryAddress, serviceCode, deliveryDODAAC,
                 orderingDODAAC, modes, notes):
        if (itemNum == ""):
            itemNum = "N/A"
        if (quantity == ""):
            quantity = "N/A"
        if (quantity8A == ""):
            quantity8A = "N/A"
        if (quantitySA == ""):
            quantitySA = "N/A"
        if (unrestricted == ""):
            unrestricted = "N/A"
        if (performancePeriod == ""):
            performancePeriod = "N/A"
        if (nsn == ""):
            nsn = "N/A"
        if (deliveryID == ""):
            deliveryID = "N/A"
        if (state == ""):
            state = "N/A"
        if (region == ""):
            region = "N/A"
        if (throughputSPLC == ""):
            throughputSPLC = "N/A"
        if (requirementSPLC == ""):
            requirementSPLC = "N/A"
        if (deliveryAddress == ""):
            deliveryAddress = "N/A"
        if (serviceCode == ""):
            serviceCode = "N/A"
        if (deliveryDODAAC == ""):
            deliveryDODAAC = "N/A"
        if (orderingDODAAC == ""):
            orderingDODAAC = "N/A"
        if (notes == ""):
            notes = "N/A"

        self.itemNum = itemNum
        self.quantity = quantity
        self.quantity8A = quantity8A
        self.quantitySA = quantitySA
        self.unrestricted = unrestricted
        self.performancePeriod = performancePeriod
        self.nsn = nsn
        self.deliveryID = deliveryID
        self.state = state
        self.region = region
        self.throughputSPLC = throughputSPLC
        self.requirementSPLC = requirementSPLC
        self.deliveryAddress = deliveryAddress
        self.serviceCode = serviceCode
        self.deliveryDODAAC = deliveryDODAAC
        self.orderingDODAAC = orderingDODAAC
        self.modes = ModeList(modes)
        self.notes = notes

    def __repr__(self):
        return f"{self.__class__.__name__}(Item Number: {self.itemNum}, Performance Period: {self.performancePeriod})"

    def __str__(self):
        return f"{self.__class__.__name__}({self.itemNum}, {self.performancePeriod})"

    def neatData(self):
        output = "" + self.itemNum + " " + self.quantity + " " + self.quantity8A + "\n"
        output += self.quantitySA + " " + self.unrestricted + " " + self.performancePeriod + " " + self.nsn + "\n"
        output += self.deliveryID + " " + self.state + " " + self.region + " " + self.throughputSPLC + "\n"
        output += self.requirementSPLC + " " + self.deliveryAddress + " " + self.serviceCode + "\n"
        output += self.deliveryDODAAC + " " + self.orderingDODAAC + " " + str(self.modes) + " " + self.notes
        output += "\n"

        return output

    def data(self):
        output = "" + self.itemNum + ", " + self.quantity + ", " + self.quantity8A
        output += ", " + self.quantitySA + ", " + self.unrestricted + ", " + self.performancePeriod + ", " + self.nsn
        output += ", " + self.deliveryID + ", " + self.state + ", " + self.region + ", " + self.throughputSPLC
        output += ", " + self.requirementSPLC + ", " + self.deliveryAddress + ", " + self.serviceCode
        output += ", " + self.deliveryDODAAC + ", " + self.orderingDODAAC + ", " + str(self.modes) + ", " + self.notes

        return output

    def list(self):
        output = [self.itemNum, self.quantity, self.quantity8A, self.quantitySA, self.unrestricted,
                  self.performancePeriod, self.nsn, self.deliveryID, self.state, self.region, self.throughputSPLC,
                  self.requirementSPLC, self.deliveryAddress, self.serviceCode, self.deliveryDODAAC,
                  self.orderingDODAAC, str(self.modes), self.notes]
        return output


class ItemList:

    def __init__(self, items=None):
        if items is None:
            items = []
        if isinstance(items, Item):
            items = [items]
        self.items = items

    def data(self):
        allItems = ""
        for item in self.items:
            allItems += item.data() + ", "
        allItems = allItems[:len(allItems) - 2]
        return f"{allItems}"

    def neatData(self):
        allItems = ""
        for item in self.items:
            allItems += item.neatData() + "\n"

        return f"({allItems}"

    def __repr__(self):
        allItems = ""
        for item in self.items:
            allItems += repr(item) + ", "
        allItems = allItems[:len(allItems) - 2]
        return f"{self.__class__.__name__}: {allItems}"

    def __str__(self):
        allItems = ""
        for item in self.items:
            allItems += str(item) + ", "
        allItems = allItems[:len(allItems) - 2]
        return f"{self.__class__.__name__}: {allItems}"


class Extractor:
    def __init__(self, filename):
        self.filename = filename

    def __repr__(self):
        return f"{self.__class__.__name__} (filename: {self.filename})"

    def __str__(self):
        return f"{self.__class__.__name__} ({self.filename})"

    def pdfDataExtraction(self):
        pdfFile = open(self.filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)

        numPages = pdfReader.numPages - 62

        everyPage = ''
        for pageNum in range(8, numPages):
            tempPage = pdfReader.getPage(pageNum)
            everyPage += tempPage.extractText()
        onlyItems = everyPage[everyPage.find("Item: 00"):]
        result = onlyItems.split("Item: ")
        result.pop(0)

        clean_result = []

        for index in range(len(result)):
            if (result[index][:2] != ' T'):
                clean_result.append(result[index])

        return clean_result

    def extractToNSN(self, rawItem):
        focus = rawItem[0:rawItem.find("NSN")]
        focuslist = focus.split()
        result = []
        # Item
        result.append(focuslist[0])
        # Quantity
        result.append(focuslist[2] + " " + focuslist[3])
        # 8A Quantity
        result.append(focuslist[6])
        # SA Quantity
        result.append(focuslist[9])
        # Unrestricted
        result.append(focuslist[11])
        # Period of Performance
        result.append(focuslist[15])

        return result

    def extractNSNtoRegion(self, rawItem):
        focus = rawItem[rawItem.find("NSN"):rawItem.find("Region")]
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

    def extractRegionToService(self, rawItem):
        result = ["Region", "ThroughputSPLC", "RequirementSPLC", "Delivery Address"]
        focus = rawItem[rawItem.find("Region"):rawItem.find("Service Code")]

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

    def extractServiceCodeToMode(self, rawItem):
        focus = rawItem[rawItem.find("Service Code"):rawItem.find("Mode")]

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

    def findMaxAndMin(self, numbers):
        result = ["max", "min"]
        secondnum = ""
        orginialnum = numbers
        index = numbers.find(",")
        if (numbers.find(",") < 0 and len(numbers) <= 3):
            firstnum = numbers
            secondnum = ""
        elif (numbers.find(",") < 0):
            firstnum = numbers[:3]
            secondnum = numbers[3:]
        else:

            while (True):

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

    def extractModeToEnd(self, rawItem):
        focus = rawItem[rawItem.find("BULK:"):]

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
                        min_max_array = self.findMaxAndMin(numbers)
                        modeinfo[2] = min_max_array[0]
                        modeinfo[3] = min_max_array[1]
                        modeinfo[4] = currentmode[
                                      endindex:endindex + 1]  # It is only the end index for the numbers slice
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
                    min_max_array = self.findMaxAndMin(numbers)

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

    def process(self):
        clean_result = self.pdfDataExtraction()
        itemlist = ItemList()
        for i in range(len(clean_result)):
            item = clean_result[i]

            itemattributes = self.extractToNSN(item)

            itemattributes += self.extractNSNtoRegion(item)

            itemattributes += self.extractRegionToService(item)

            itemattributes += self.extractServiceCodeToMode(item)

            itemattributes.append(self.extractModeToEnd(item))

            itemObj = Item(itemattributes[0], itemattributes[1], itemattributes[2], itemattributes[3],
                           itemattributes[4], itemattributes[5], itemattributes[6], itemattributes[7],
                           itemattributes[8], itemattributes[9], itemattributes[10], itemattributes[11],
                           itemattributes[12], itemattributes[13], itemattributes[14], itemattributes[15],
                           itemattributes[16][:len(itemattributes[16]) - 1],
                           itemattributes[16][len(itemattributes[16]) - 1:])
            itemlist.items.append(itemObj)

        workbook = xlsxwriter.Workbook('my.csv')
        worksheet = workbook.add_worksheet('my.csv')

        # Writes Data into sheet
        for i in range(len(itemlist.items)):
            for j in range(len(itemlist.items[i].list())):
                worksheet.write(i + 1, j, str(itemlist.items[i].list()[j]))

        itemFormat = ["Item#", "Quantity", "8A Quantity", "SA Quantity", "Unrestricted", "Period of Performance",
                      "NSN", "Delivery Identification", "State", "Region", "Throughput SPLC", "Requirement SPLC",
                      "Delivery Address", "Service Code", "Delivery DODAAC", "Ordering Office DODAAC",
                      "Modes", "Delivery Notes"]
        for i in range(len(itemFormat)):
            worksheet.write(0, i, itemFormat[i])

        workbook.close()


def main():
    #   replace "filename" variable with the name of the fuel tender document you have
    filename = 'WESTPAC_sol.pdf'
    extractor = Extractor(filename)
    extractor.process()


if __name__ == '__main__':
    main()
