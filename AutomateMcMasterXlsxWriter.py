import pyperclip, re, openpyxl, os, xlsxwriter
from openpyxl.styles import Font, Style, Alignment, Border, Side, Color

text = str(pyperclip.paste())

itemNumberRegex = re.compile(r'[0-9A-Z]{6,9}')  # item serial number regex pattern
itemQuantityRegex = re.compile(r'[0-9]{1,4}')  # item quantity regex pattern
itemPriceRegex = re.compile(r'[0-9]{1,4}.[0-9]{1,4}')  # item price regex pattern
itemDescriptionRegex = re.compile(r'.*',re.DOTALL)  # item description regex pattern
itemLineNumberRegexInit = re.compile(r'\d\d? \r\n')  # initial item line number regex pattern (only used as reference)
itemLineNumberRegex = re.compile(r'\r\n\d\d? \r\n')  # item line number regex pattern (only used as reference)

mo = itemNumberRegex.findall(text)  # find number of items based on quantity of serial number regex pattern
numberOfItems = len(mo)

if numberOfItems is 0:  # checks if copied text is valid
    print('Error: No items found. Are you sure you have the correct text copied onto the clipboard?')
    # print('Press ENTER to try again. Enter \'q\' to quit')
    # input = input()
    # if 'q' in input:
    #     break
    # elif '' in input:
    #     continue
    # else:
    #     print('Invalid input.')
# elif numberOfItems > 15:
#     print('Too many items to fit in one spreadsheet. Consider splitting into two or more spreadsheets.')
else:

    # initialization of lists and variables
    itemNumbersStr = []
    itemQuantitiesStr = []
    itemPricesStr = []
    itemDescriptionsStr = []
    itemQuantitiesNum = []
    itemPricesNum = []
    itemTotalPricesNum = []
    itemTotalPricesStr = []
    totalPrice = 0
    startPos = 0

    for x in range(0, numberOfItems):
        # searches for item serial number with regex pattern using defined starting position
        itemNumberSearch = itemNumberRegex.search(text, startPos)
        itemNumbersStr.append(itemNumberSearch.group())  # appends matched serial number to list

        # update the starting position counter for next loop iteration
        startPos = itemNumberSearch.span()[1]

        # searches for item quantity number with regex pattern using end position from serial number match as reference
        itemQuantitySearch = itemQuantityRegex.search(text, itemNumberSearch.span()[1])
        itemQuantitiesStr.append(itemQuantitySearch.group())  # appends matched quantity number to list as string value
        # convert quantity list from strings to integers
        itemQuantitiesNum.append(int(itemQuantitiesStr[x]))


        # searches for item quantity number with regex pattern using end position from quantity match as reference
        itemPriceSearch = itemPriceRegex.search(text, itemQuantitySearch.span()[1])
        itemPricesStr.append(itemPriceSearch.group())  # appends matched price number to list as string value
        # convert price list from strings to floats
        itemPricesNum.append(float(itemPricesStr[x]))

        # calculate total prices for products
        product = round(itemQuantitiesNum[x] * itemPricesNum[x], 2)
        itemTotalPricesNum.append(product)
        # convert total price floats to string
        itemTotalPricesStr.append(str(itemTotalPricesNum[x]))
        totalPrice += product  # compute total price

        # search for line number at start of text using position from item number match as reference
        if x is 0:
            itemLineNumberSearch = itemLineNumberRegexInit.search(text, 0, itemNumberSearch.span()[0])
            itemDescriptionSearch = itemDescriptionRegex.search(text, itemLineNumberSearch.span()[1], itemNumberSearch.span()[0])

        # search for line number for rest of loop iteration using position from price match as reference
        elif 0 < x < numberOfItems:
            itemLineNumberSearch = itemLineNumberRegex.search(text, itemLineNumberSearchNextStartPos, itemNumberSearch.span()[0])
            itemDescriptionSearch = itemDescriptionRegex.search(text, itemLineNumberSearch.span()[1], itemNumberSearch.span()[0])

        # update the starting position counter for line number search for next loop iteration
        itemLineNumberSearchNextStartPos = itemPriceSearch.span()[1]

        itemDescriptionsStr.append(itemDescriptionSearch.group())  # appends matched description to list

        # clean up item description
        itemDescriptionsStr[x] = itemDescriptionsStr[x].replace(' \r\n', '')
        itemDescriptionsStr[x] = itemDescriptionsStr[x].replace('\r\n', ' ')
        itemDescriptionsStr[x] = itemDescriptionsStr[x].lstrip(' ')
        itemDescriptionsStr[x] = itemDescriptionsStr[x].rstrip(' ')

        # get rid of unnecessary spaces before commas
        itemDescriptionsStr[x] = re.sub(r'( +,)', ',', itemDescriptionsStr[x])

    print('Successfully converted items. Writing to excel file.')

    wb = xlsxwriter.Workbook('purchaseorder.xlsx')

    # sheet = wb.get_sheet_by_name('Purchase Order')
    #
    # customFont = Font(name='Roboto', size=8, bold=False, underline='none')
    # customAlignment = Alignment(vertical='center', wrapText=True)
    # customBorder = Border(left=Side(style='thin', color=Color(rgb='FF3B5E91')),
    #                       right=Side(style='thin', color=Color(rgb='FF3B5E91')),
    #                       top=Side(style='thin', color=Color(rgb='FF3B5E91')),
    #                       bottom=Side(style='thin', color=Color(rgb='FF3B5E91')))
    #
    # styleObj = Style(font=customFont, alignment=customAlignment, border=customBorder)
    #
    # spreadsheetsRequired = (numberOfItems + 1) // 16
    #
    # for x in range(20, 36):
    #     sheet['A' + str(x)].style = styleObj
    #     sheet['B' + str(x)].style = styleObj
    #     sheet['C' + str(x)].style = styleObj
    #     sheet['G' + str(x)].style = styleObj
    #     sheet.merge_cells('C' + str(x) + ':' + 'D' + str(x))
    #
    # for x in range(0, numberOfItems):
    #     sheet['A' + str(x+20)] = itemQuantitiesStr[x]
    #     sheet['B' + str(x+20)] = itemNumbersStr[x]
    #     sheet['C' + str(x+20)] = itemDescriptionsStr[x]
    #     if len(itemDescriptionsStr[x]) > 30:
    #         heightFactor = len(itemDescriptionsStr[x]) // 30
    #         sheet.row_dimensions[x+20].height = 12 * heightFactor
    #     else:
    #         sheet.row_dimensions[x+20].height = 12
    #     sheet['G' + str(x+20)] = itemPricesStr[x]

    #wb.save('purchaseorder.xlsx')

