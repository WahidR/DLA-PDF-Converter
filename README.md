# DLA-PDF-Converter
Converts DLA/DOD contract PDFs to Excel

- Utilizes "PyPDF2"



- mainlist is a list of item Arrays

- This is how each item array will look
    -itemArrayFormat = ["Item#", "Quantity", "8A Quantity", "SA Quantity", "Unrestricted", "Period of Performance",
    "NSN", "Delivery Identification", "State", "Region", "Throughput SPLC", "Requirement SPLC",
    "Delivery Address","Service Code", "Delivery DODAAC", "Ordering Office DODAAC", "Notes"]


- modesFormat is another array because there can be multiple modes in each item. Each mode will be in a mode list
    -modesFormat = ["Mode", "FOB Restriction", "FSII", "SDA", "CI"]

- mode list will be [mode1, mode2, mode3] or just [mode1]

- Delivery Notes and Delivery Hours for each item will be one single string called "notes"
    -notes = "Delivery Hours: 7AM - 12PM Delivery Notes: blablabla"
    -I haven't decided as to where I should add notes. The modes will be a list so I could add notes to the end of the modes list or I could add notes to the end of itemArrayFormat


