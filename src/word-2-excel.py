import csv
from itertools import permutations
from datetime import datetime
import sys

import docx2txt

from convert_am_pm import convert24


#
# Create a simple .csv file
#

def create_excel_file(filename, sortedArray):

    filename = filename.replace('docx','csv')
    with open(filename, mode='w', newline='') as _file:
        _writer = csv.writer(_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        _writer.writerow(['From Status', 'To Status', 'Effective Time'])

        for _cell in sortedArray:
            _writer.writerow([_cell[1], _cell[2], _cell[3]])


#
# Make a list of all possible 2 combo keywords
#

def make_permutations():
    results = []

    keywords = ['Change Request', 'Suspended', 'In Progress', 'Acknowledged', 'Dispatched', 'Opened']

    # Get all permutations of length 2
    perm = permutations(keywords, 2)

    # Make the possible clauses
    for _list in list(perm):
        search_term = "from " + _list[0] + " to " + _list[1]
        results.append(search_term)

    # Hard code in a 'to Opened'.  This is the starting state.
    results.append('to Opened')
    return results


def find_clause_index(my_text, substring, start):
    nlf = '\n'
    _idx = my_text.find(substring, start)
    result = ""

    # We found a clause
    if _idx != -1:

        # Find the next newline character starting at the index we found our clause at
        index_of_nextline = my_text.find(nlf, _idx)

        # Find the one after that. We stop when we find a line that contains AM and/or PM

        not_found = True
        while not_found:
            result = my_text[_idx:index_of_nextline]
            if result.find("AM") != -1 or result.find("PM") != -1:
                not_found = False
            else:
                _idx = index_of_nextline
                index_of_nextline = my_text.find(nlf, _idx + 2)

    return _idx, result


if __name__ == "__main__":

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
        my_text = docx2txt.process(filename)
    except:
        print("Error extracting text from document: " + filename)    

    for perm in perms:
        not_eof = True
        previous_index = 0
        while not_eof:
            idx, snippet = find_clause_index(my_text, perm, previous_index)
            if idx != -1:
                # This perm was found in the file.  We will store this index & process this entry
                # remembering that this perm may occur multiple times in the file.
                previous_index = idx

                idx = snippet.find('(effective')
                if idx != -1:
                    augment = len('(effective')
                    paren = snippet.find(')')
                    snippet = snippet[idx+augment:paren].strip()

                if 'PM' in snippet:
                    snippet = snippet.replace(' PM', ':00 PM')

                if 'AM' in snippet:
                    snippet = snippet.replace(' AM', ':00 AM')

                _snippet_list = snippet.split(" ", 1)
                tmp = convert24(_snippet_list[1].strip())
                _snippet_list[1] = tmp

                string_date = ' '.join(_snippet_list)
                string_date = string_date.replace('\n', '')
                string_date = string_date.strip()

                # Convert String ( ‘DD/MM/YY HH:MM:SS ‘) to datetime object
                datetime_obj = datetime.strptime(string_date, '%m/%d/%y %H:%M:%S')

                perm = perm.strip()

                if "from " not in perm:
                    frompart = "None"
                    topart = perm[perm.index("to ") + 3:]
                else:
                    frompart = perm[5:perm.index(" to ")]
                    topart = perm[perm.index(" to ") + 4:]
                
                d = [datetime_obj, str(frompart), str(topart), string_date]
                final_list.append(d)

            else:
                not_eof = False

    sortedArray = sorted(
        final_list,
        key=lambda x: datetime.strptime(x[3], '%m/%d/%y %H:%M:%S'), reverse=False
    )

    create_excel_file(filename, sortedArray)
    print("\nRun completed.")
