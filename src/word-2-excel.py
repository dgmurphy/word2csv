import csv
from itertools import permutations
import re
import sys
from zipfile import ZipFile
import datetime

import lxml.etree as ET


# Set constants.
KEYWORDS = [
            'Opened',
            'Acknowledged',
            'Dispatched',
            'In Progress',
            'Change Request',
            'Suspended',
            'Repaired',
            'Completed',
            'Closed',
            'Rejected',
            'Cancelled',
            'Reopened'
           ]
time_format = "%m/%d/%y %I:%M %p"


def create_excel_file(filename, sorted_list):
    """
    Create a simple .csv file

    """

    filename = filename.replace('docx','csv')
    with open(filename, mode='w', newline='') as _file:
        _writer = csv.writer(_file, delimiter=',', quotechar='"',
                             quoting=csv.QUOTE_MINIMAL)
        _writer.writerow(['From Status', 'To Status', 'Effective Time'])

        for _cell in sorted_list:
            _writer.writerow([_cell[0], _cell[1],
                              _cell[2].strftime(time_format)])


def make_permutations():
    """
    Make a list of all possible 2 combo keywords

    """

    results = []

    # Get all permutations of length 2
    perms = permutations(KEYWORDS, 2)

    # Make the possible clauses
    for _list in list(perms):
        search_term = "from " + _list[0] + " to " + _list[1]
        results.append(search_term)

    # Hard code in a 'to Opened'.  This is the starting state.
    results.append('to Opened')

    return results


def lineStartsWithADate(line):

    line = line.strip()

    match = re.search("^(^\d\d\/\d\d\/\d\d \d\d:\d\d .M)",line)
    if match is not None:
        return True
    else:
        return False



if __name__ == "__main__":


    # Get the status change permutations.
    perms = make_permutations()

    # Attempt to find these substrings in our file.
    # We are going to start at the beginning of the file.
    # We could start at the index of "Updates" but we will not bother.
    # A clause can occur multiple times within the history of action.

    if len(sys.argv) < 2:
        print("ERROR: No input file specified.")
        sys.exit()

    filename = sys.argv[1]
    final_list = []
    date_list = []

    # Try getting the text
    try:
        word = ZipFile(filename)
    except Exception as e:
        print("Error opening Word document: " + filename)
        print(e)
        raise

    # Read the document body into an XML Element object, and store the
    # document namespace map.
    document_root= ET.fromstring(word.read("word/document.xml"))
    ns = document_root.nsmap
    body = document_root.find("w:body", namespaces = ns)

    # Find the title for the "Updates" table, then search through the table
    # itself (which will be the next child element) and iterate through its
    # rows to locate status changes.

    updates_found = False
    for child in body:

        # Check to see if the previous child was the "Updates" title. If it
        # is, search the table for updates, and then break the loop.

        if updates_found:

            # Find all table cells that match each of the status change
            # permutations.

            for perm in perms:

                # Find all matching cells (if any).
                query = './/w:tc[descendant::*[contains(text(), "' +  \
                        perm + '")]]'
                matches = child.xpath(query, namespaces = ns)

                # If there are no matches--which will be true in the vast
                # majority of cases--this will be an empty loop.
                for match in matches:

                    # The time is always the last paragraph in an update
                    # cell. However, this one-line paragraph is sometimes
                    # split across two (or possibly more) text elements.
                    time_para = match[-1].findall(".//w:t", namespaces = ns)
                    time_para_string = "".join([string.text for string
                                                in time_para])

                    # Pull out the effective time, discarding the time of
                    # update entry, if any (there's usually an entry time,
                    # but not always--hence we pull out the last sub-
                    # string).
                    time_string = \
                        time_para_string.split("effective")[-1].strip(" )")
                    time = datetime.datetime.strptime(time_string, time_format)

                    if perm.startswith("to "):
                        from_part = "None"
                        to_part = perm.split("to ")[1]
                    else:
                        perm_split = perm.split(" to ")
                        from_part = perm_split[0].split("from ")[1]
                        to_part = perm_split[1]

                    d = [from_part, to_part, time]
                    final_list.append(d)

            # If we've found and searched the "Updates" table, we're done.
            break

        # Unless the "Updates" title is found, this will be an empty list,
        # which will evaluate as False in the Boolean test at the beginning
        # of the loop.
        updates_found = child.xpath('.//w:t[text()="Updates"]', namespaces = ns)

    sorted_list = sorted(final_list, key = lambda x: x[2], reverse = False)

    create_excel_file(filename, sorted_list)
    print("\nRun completed.")
