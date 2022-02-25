import ezsheets

ss = ezsheets.Spreadsheet('Google-Sheets-URL-code')

sheet = ss['Parser List']

def parens(cell):
    parens = []
    step = 0
    for i in cell:
        if i == '(':
            parens += [step]
            step += 1
        elif i == ')':
            parens += [step]
            step += 1
        else:
            step += 1
    return (parens)

def cellQuantities(cell):
    listOfParens = parens(cell)
    quant = []
    uniqueItems = int(((len(listOfParens))/2))
    start = 0
    stop = 1
    if ((len(listOfParens)) % 2) == 0:
        for i in range(uniqueItems):
            if (listOfParens[start]+1) == (listOfParens[stop]-1):
                quant.append(cell[(listOfParens[start]+1)])
                start += 2
                stop += 2
            else:
                quant.append(cell[(listOfParens[start]+1):(listOfParens[stop])])
                start += 2
                stop += 2
        return (quant)
    else:
        return ('error - Information format is incorrect')

def cellItems(cell):
    listOfParens = parens(cell)
    quant = []
    uniqueItems = int(((len(listOfParens))/2))
    start = 1
    stop = 2
    if ((len(listOfParens)) % 2) == 0:
        for i in range(uniqueItems):
            if i == (uniqueItems-1):
                quant.append(cell[(listOfParens[start])+1:])
            else:
                quant.append(cell[(listOfParens[start])+1:(listOfParens[stop])])
                start += 2
                stop += 2
        return (quant)
    else:
        return ('error - Information format is incorrect')

def fastenerColumnSet(columnNumber):
    fastenerColumn = sheet.getColumn(columnNumber)
    row = 1
    fastenerList = []
    for i in fastenerColumn:
        fastenerList += (cellItems((sheet [columnNumber, row])))
        row += 1
    fastenerSet = set(fastenerList)
    return (fastenerSet)

def findFastenerColumn(sheet):
    column = 1
    for i in sheet.getRow(1):
        if i != 'Information':
            column += 1
        elif i == '':
            print ('error - could not find Information Column')
            break
        else:
            break
    return (column)

def findQtyColumn(sheet):
    column = 1
    for i in sheet.getRow(1):
        if i != 'Qty':
            column += 1
        elif i == '':
            print ('error - could not find Information Column')
            break
        else:
            break
    return (column)

def fastenerCountColumns(sheet):
    fastenerCulumnNumber = findFastenerColumn(sheet)
    fastenerList = list(fastenerColumnSet(fastenerCulumnNumber))
    fastenerList.sort()
    column = 1
    row = 1
    for i in sheet.getRow(row):
        if (sheet [column, row]) != '':
            column += 1
        else:
            break
    for i in fastenerList:
        print(i)
        (sheet [column, row]) = i
        column += 1

def propagateRowQuantity(row):
    column = findFastenerColumn(sheet)
    cell = sheet [column, row]
    qtyColumn = findQtyColumn(sheet)
    qty = sheet [qtyColumn, row]
    fastenerQtyList = cellQuantities(cell)
    fastenerItemList = cellItems(cell)
    placeMarker = 0
    if qty != '':
        for item in fastenerItemList:
            totalItemQty = (int(qty) * int(fastenerQtyList[placeMarker]))
            placeMarker += 1
            destinationMarker = 1
            for i in sheet.getRow(1):
                if item == i:
                    sheet [destinationMarker, row] = totalItemQty
                    break
                else:
                    destinationMarker += 1
    else:
        return()

def propagateAllItemQuantities(sheet):
    column = findFastenerColumn(sheet)
    row = 1
    for i in sheet.getColumn(column):
        if row == 1:
            row += 1
            print(row)
        else:
            propagateRowQuantity(row)
            row += 1
            print(row)

def runIt(sheet):
    sheet.refresh()
    fastenerCountColumns(sheet)
    propagateAllItemQuantities(sheet)
