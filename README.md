# DLA-PDF-Converter
Converts DLA/DOD contract PDFs to Excel


# mainlist is a list of item Arrays

# itemsArrayFormat shows the format of how EACH item array will look
itemArrayFormat = ["Item#", "Quantity", "8A Quantity", "SA Quantity", "Unrestricted", "Period of Performance",
"NSN", "Delivery Identification", "State", "Region", "Throughput SPLC", "Requirement SPLC",
"Delivery Address","Service Code", "Delivery DODAAC", "Ordering Office DODAAC", "Notes"]


# modesFormat is another array because there can be multiple modes in each item. Each mode will be in a mode list
modesFormat = ["Mode", "FOB Restriction", "FSII", "SDA", "CI"]

# mode list will be [mode1, mode2, notes] or [mode1, notes]

# Delivery Notes and Delivery Hours for each item will be one single string called "notes"
notes = "Delivery Hours: 7AM - 12PM Delivery Notes: blablabla"

