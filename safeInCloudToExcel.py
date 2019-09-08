import argparse
import openpyxl
import sys
import xml.etree.ElementTree as ET


def parseArgs(args):
    '''
    Parse the command line arguments into a dictionary object.

    args [in] A list of command line arguments.
    returns   A dictionary of the parse results.
    '''

    parser = argparse.ArgumentParser(description = 'Create an Excel workbook from the contents of an '
                                     'exported SafeInCloud XML file.')

    parser.add_argument('infile',  metavar = '<input XML file>',
                        help = 'The exported SafeInCloud XML file to read.')
    parser.add_argument('outfile', metavar = '<output Excel file>',
                        help = 'The Excel .xlsx file to write to.')

    parsedargs = parser.parse_args(args)

    argDict = vars(parsedargs)

    return argDict

def convert(infile, outfile):
    # Open the XML file and get the root element.
    #
    tree     = ET.parse(infile)
    database = tree.getroot()

    # Iterate through the elements and collect entries.
    #
    elementList = list()
    fieldSet = set(('Title', 'Notes'))

    for child in database:
        # If the child object is not a card or a note, move on.
        #
        if child.tag not in ('card', 'note'):
            continue

        # If the child is deleted or a template, move on.
        #
        if 'true' == child.attrib.get('deleted', 'false'):
            continue

        if 'true' == child.attrib.get('template', 'false'):
            continue

        # Get the title for the entry. If there isn't one, move on.
        #
        title = child.attrib.get('title', None)

        if title is None:
            continue

        # Build the dictionary of entry fields, starting with the title.
        #
        elementDict = { 'Title' : title }

        for field in child:
            # If the field has no contents, move on.
            #
            if field.text is None or '' == field.text:
                continue

            # If the field is a regular field or a note, process it. If not,
            # move on.
            #
            if 'field' == field.tag:
                # Get the name of the field. If there isn't a name, move on.
                #
                name = field.attrib.get('name', None)

                if name is None:
                    continue

                # Add the name to the set of field names.
                #
                fieldSet.add(name)

            elif 'notes' == field.tag:
                # This is a "notes" field, so set the name to "Notes".
                #
                name = 'Notes'
            else:
                continue

            # Add the field to the dictionary keyed by its name.
            #
            elementDict[name] = field.text

        # Add the element to the list.
        #
        elementList.append(elementDict)

    # Create a field list with Title, Login, and Password first, Notes
    # last, and all other fields alphabetically ordered in the middle.
    #
    fieldList = [ 'Title', 'Login', 'Password' ]

    for name in fieldList + [ 'Notes' ]:
        try:
            fieldSet.remove(name)
        except:
            pass

    subFieldList = list(fieldSet)

    subFieldList.sort()

    fieldList.extend(subFieldList)

    fieldList.append('Notes')

    # Create the output workbook and get a sheet.
    #
    book = openpyxl.Workbook()
    sheet = book.active

    # Add the list of fields as a header.
    #
    sheet.append(fieldList)

    # Iterate through the elements and add a row for each.
    #
    for elementDict in elementList:
        row = list()

        for field in fieldList:
            row.append(elementDict.get(field, None))

        sheet.append(row)

    # Save the workbook.
    #
    book.save(filename = outfile)


if __name__ == '__main__':
    # Get the arguments.
    #
    argDict = parseArgs(sys.argv[1:])

    # Convert the input XML into an Excel spreadsheet.
    #
    convert(**argDict)
