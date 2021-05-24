import docx2txt
from convert_am_pm import convert24
from datetime import datetime

#
# Make a list of all possible 2 combo keywords
#

def make_permutations():
    results = []
    from itertools import permutations

    keywords = ['Change Request', 'Suspended', 'In Progress', 'Acknowledged', 'Dispatched', 'Opened']

    # Get all permutations of length 2
    perm = permutations(keywords, 2)

    # Make the possible clauses
    for _list in list(perm):
        search_term = "from " + _list[0] + " to " + _list[1]
        results.append(search_term)
    return results


def find_clause_index(_filename, substring, start):
    nlf = '\n'
    my_text = docx2txt.process(_filename)
    _idx = my_text.find(substring,start)
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

    filename = "./files/42-SEYM-200728-0002_2021-18-05_09-58-AM.docx"
    final_list = []

    for perm in perms:
        not_eof = True
        previous_index = 0
        while not_eof:
            idx, snippet = find_clause_index(filename, perm, previous_index)
            if idx != -1:
                # This perm was found in the file.  We will store this index & process this entry
                # remembering that this perm may occur multiple times in the file.
                previous_index =  idx

                idx = snippet.find('(effective')
                if idx != -1:
                    snippet = snippet[0:idx]

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
                print(string_date)
                print(str(perm))

                # Convert String ( ‘DD/MM/YY HH:MM:SS ‘) to datetime object
                datetime_obj = datetime.strptime(string_date, '%m/%d/%y %H:%M:%S')
                d = [{'date-object' : datetime_obj},
                    {'clause': str(perm)},
                    {'date-string': string_date}]
                final_list.append(d)
            else:
                not_eof = False

    # sortedArray = sorted(
    #     final_list,
    #     key=lambda x: datetime.strptime(x['date-object'], '%m/%d/%y %H:%M:%S'), reverse=True
    # )

    # print (sortedArray)