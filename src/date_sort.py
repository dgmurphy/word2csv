# Python3 program to sort the list of dates given in string format

# Import the datetime module
from datetime import datetime


# TODO ---------------------------------------------------------------
# TODO - Needs completion.
# TODO ---------------------------------------------------------------

# Function to print the data stored in the list
def printDates(dates):
    for i in range(len(dates)):
        print(dates[i])


if __name__ == "__main__":
    dates = ["23 Jun 2018", "2 Dec 2017", "11 Jun 2018",
             "01 Jan 2019", "10 Jul 2016", "01 Jan 2007"]

    # Sort the list in ascending order of dates
    dates.sort(key=lambda date: datetime.strptime(date, '%d %b %Y'))

    # Print the dates in a sorted order
    printDates(dates)

    dates = ["02/09/28", "07/28/20", "05/12/21",
             "05/05/31", "06/12/00", "11/11/11"]

    dates = ["02 09 28", "07 28 20", "05 12 21",
             "05 05 31", "06 12 00", "11 11 11"]

    # Sort the list in ascending order of dates
    dates.sort(key=lambda date: datetime.strptime(date, '%d %b %Y'))

    # Print the dates in a sorted order
    printDates(dates)
