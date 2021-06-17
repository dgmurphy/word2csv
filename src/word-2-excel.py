import csv
from sys import argv, stdout
import sys
from time import sleep
import re
import datetime # We avoid direct imports of classes with confusing names.
from datetime import timedelta
from zipfile import ZipFile
from os.path import splitext
from io import BytesIO

import holidays
import lxml.etree as ET


# Set constants. To make it easier to deploy the app, as a single file,
# everything is hard-coded here, but to facilitate changes, everything is
# a constant, set here at the beginning of the script.
DETAILS_TITLE = "Details"
TICKET_LABEL = "Case Number"
CONTRACT_LABEL = "Contract"
CONTRACT_PREFIXES = ["USAF -", "FSS-III CONUS -"]
PRIORITY_LABEL = "Priority"
UPDATES_TITLE = "Updates"
STATUS_CHANGE_PHRASE = "changed the status of"
UPDATER_SEP = " ("
TO_SEP = " to "
STATUS_NOTE_PATT_STRING = r" \((.+)\)" # regex for " (<status note>)"
FROM_SEP = " from "
EFFECTIVE_TIME_SEP = " (effective "
EFFECTIVE_TIME_STRIP_CHARS = ")"
REPORT_TIME_PREFIX = "Generated: "
TIME_FORMAT = "%m/%d/%y %I:%M %p"
MTRF_RULES = {"PL1: Catastrophic Failure": {"hours": 12, "duty day": False},
              "PL1: Major Malfunction": {"hours": 12, "duty day": False},
              "PL1: Partial Failure": {"hours": 24, "duty day": True},
              "PL2: Catastrophic Failure": {"hours": 12, "duty day": False},
              "PL2: Major Malfunction": {"hours": 12, "duty day": False},
              "PL2: Partial Failure": {"hours": 24, "duty day": True},
              "PL3: Catastrophic Failure": {"hours": 12, "duty day": False},
              "PL3: Major Malfunction": {"hours": 12, "duty day": True},
              "PL3: Partial Failure": {"hours": 24, "duty day": True},
              "PL4: Catastrophic Failure": {"hours": 24, "duty day": False},
              "PL4: Major Malfunction": {"hours": 24, "duty day": True},
              "PL4: Partial Failure": {"hours": 24, "duty day": True}}
DUTY_DAY = {"start": datetime.time(hour = 7), "end": datetime.time(hour = 16)}
HOLIDAYS = holidays.UnitedStates()
SUSPEND_STATUSES = ["Suspended"]


def calculate_delay(start, end, priority):
    """
    Note that, at present, we don't use the "hours" values from
    "MTRF_RULES"--we're simply calculating the times between status
    changes, not comparing them to a metric.

    """

    # If we have to include weekends, holidays, and time outside the duty
    # day in the delay between status changes, our calculation is simple.
    if not MTRF_RULES[priority]["duty day"]:
        delay = end - start

    # If we don't have to count weekends, holidays, and time outside the
    # duty day, things get much more complex.

    else:

        # First, we deal with the start date of the status period. If the
        # start date was on a weekend or holiday, or the start time was
        # after the end of the duty day, we don't record any time for
        # this day.
        if start.time() > DUTY_DAY["end"] or start.isoweekday() >= 6 or \
          start in HOLIDAYS:
            delay = timedelta(0)

        # Otherwise, we record only time during the duty day.
        else:

            # If the period started before the duty day, reset "start"
            # to the beginning of the duty day.
            if start.time() < DUTY_DAY["start"]:
                start = datetime.datetime.combine(start.date(),
                                                  DUTY_DAY["start"])

            # If the end of the status period was on the same day and
            # occured before the end of the duty day, we subtract the start
            # time from that time to get our delay--unless both the end
            # update and the start update occurred before the start of the
            # duty day, in which case "end - start" calculation will
            # produce a negative number, which we'll correct to 0.
            if end.date() == start.date() and end.time() < DUTY_DAY["end"]:
                delay = max(end - start, timedelta(0))

            # If the end update occurred after the duty day, we use the
            # end of the duty day to calculate time for the first day.
            else:
                delay = \
                    datetime.datetime.combine(start.date(), DUTY_DAY["end"]) - \
                    start

            # of the duty day

        # If the end of the status period was on a different day than the
        # start, we need to add time for any additional day(s).
        if end.date() > start.date():

            # In many cases, we'll need to know the length of a duty day--
            # this is slightly complex because we can't perform arithmetic
            # on the "datetime.time" class.
            arb_date = datetime.date(1, 1, 1)
            duty_time = datetime.datetime.combine(arb_date, DUTY_DAY["end"]) - \
                        datetime.datetime.combine(arb_date, DUTY_DAY["start"])

            # We'll only add time for the end date if the end date was not
            # on a weekend or holiday, and the end time was not before the
            # beginning of the duty day.
            if end.time() > DUTY_DAY["start"] and start.isoweekday() <= 5 and \
              end not in HOLIDAYS:

                # We add the difference between the beginning of the duty
                # day and the end time, or, if the end occurs after the
                # duty day, we add the length of the duty day.
                if end.time() > DUTY_DAY["end"]:
                    delay += duty_time
                else:
                    delay += end - datetime.datetime.combine(end.date(),
                                                             DUTY_DAY["start"])

            # If the end date was at least two days after the start date
            # (meaning there's at least one full day in between), we need
            # to account for the days between the start and end updates.
            if end.date() >= (start + timedelta(days = 2)).date():

                # For each day that's not a weekend or holiday, wee add the
                # entire duty day to the total.
                current_full_day = start.date() + timedelta(days = 1)
                while current_full_day < end.date():
                    if current_full_day.isoweekday() <= 5 and \
                      current_full_day not in HOLIDAYS:
                        delay += duty_time
                    current_full_day += timedelta(days = 1)

    # Express the delay in hours.
    delay_hours = delay.total_seconds() / 3600

    return delay_hours


def write_single_ticket_CSV(filename, details, updates_list):
    """
    Create a CSV file with details and updates for a single ticket.

    """

    with open(filename, mode='w', newline='') as _file:
        _writer = csv.writer(_file, delimiter=',', quotechar='"',
                             quoting=csv.QUOTE_MINIMAL)
        _writer.writerow(["Ticket No.", details["ticket"]])
        _writer.writerow(["Site", details["site"]])
        _writer.writerow(["Priority", details["priority"]])
        _writer.writerow(["Report Time", details["report time"]])
        _writer.writerow([])
        _writer.writerow([])
        _writer.writerow(['Entered By', 'Entered On', 'From Status',
                          'To Status', 'Status Note', 'Effective Time',
                          'Status Hours', 'Update Delay'])
        for _cell in updates_list:
            _writer.writerow((_cell["updater"],
                              _cell["entry time"].strftime(TIME_FORMAT),
                              _cell["from status"], _cell["to status"],
                              _cell["status note"],
                              _cell["effective time"].strftime(TIME_FORMAT),
                              _cell["status hours"], _cell["update delay"]))


def write_details_CSV(filename, details_list):
    """
    Create a CSV file with details from all the tickets in a multi-ticket
    ZIP file.

    """

    with open(filename, mode='w', newline='') as _file:
        _writer = csv.writer(_file, delimiter=',', quotechar='"',
                             quoting=csv.QUOTE_MINIMAL)
        _writer.writerow(['Ticket No.', 'Site', 'Priority', 'Report Date'])
        for _cell in details_list:
            _writer.writerow((_cell["ticket"], _cell["site"], _cell["priority"],
                              _cell["report time"].strftime(TIME_FORMAT)))


def write_updates_CSV(filename, all_updates):
    """
    Create a CSV file with updates from all the tickets in a multi-ticket
    ZIP file.

    """

    with open(filename, mode='w', newline='') as _file:
        _writer = csv.writer(_file, delimiter=',', quotechar='"',
                             quoting=csv.QUOTE_MINIMAL)
        _writer.writerow(['Ticket No.', 'Entered By', 'Entered On',
                          'From Status', 'To Status', 'Status Note',
                          'Effective Time', 'Status Hours', 'Update Delay'])
        for _cell in all_updates:
            _writer.writerow((_cell["ticket"], _cell["updater"],
                              _cell["entry time"].strftime(TIME_FORMAT),
                              _cell["from status"], _cell["to status"],
                              _cell["status note"],
                              _cell["effective time"].strftime(TIME_FORMAT),
                              _cell["status hours"], _cell["update delay"]))


def parse_ticket(word_file):

    # Read the document body and footer into XML Element objects, and
    # store the document and footer namespace maps.
    document_root= ET.fromstring(word_file.read("word/document.xml"))
    doc_ns = document_root.nsmap
    body = document_root.find("w:body", namespaces = doc_ns)
    footer= ET.fromstring(word_file.read("word/footer1.xml"))
    footer_ns = footer.nsmap

    # We need to find the "Details" and "Updates" tables--we'll iterate
    # through the children of the document body, and use the table titles
    # to locate the tables we need--each table should be the next child
    # after the child containing its title.

    # Initialize the flags that tell us whether we've found the table
    # titles.
    details_found = False
    updates_found = False

    # Initialize the updates list.
    updates_list = []

    # Compile a regex to detect parenthetical status notes.
    status_note_patt = re.compile(STATUS_NOTE_PATT_STRING)

    for child in body:


        # If the previous child was the "Details" title, pull the values out
        # of the "Details" table.

        if details_found:

            # Get the ticket number. Just in case the ticket number gets
            # split over multiple text elements, we join all elements
            # found into a single string.
            ticket_query = './/w:tr[descendant::*[text() = "' + TICKET_LABEL + \
                           '"]]'
            ticket_row = child.xpath(ticket_query, namespaces = doc_ns)
            ticket_cells = ticket_row[0].xpath(".//w:tc", namespaces = doc_ns)
            ticket_para = ticket_cells[1].xpath(".//w:t", namespaces = doc_ns)
            ticket = "".join([string.text for string in ticket_para])

            # Get the name of the contract--we'll use this to get the site.
            contract_query = './/w:tr[descendant::*[text() = "' + \
                             CONTRACT_LABEL + '"]]'
            contract_row = child.xpath(contract_query, namespaces = doc_ns)
            contract_cells = contract_row[0].xpath(".//w:tc",
                                                   namespaces = doc_ns)
            contract_para = contract_cells[1].xpath(".//w:t",
                                                    namespaces = doc_ns)
            contract_string = "".join([string.text for string in contract_para])

            # Pull the site name out of the contract string.
            site = None
            for prefix in CONTRACT_PREFIXES:
                if contract_string.startswith(prefix):
                    site = contract_string.split(prefix)[-1]
                    break
            if not site:
                raise ValueError("Unknown contract name format")

            # Get the ticket priority.
            priority_query = './/w:tr[descendant::*[text() = "' + \
                             PRIORITY_LABEL + '"]]'
            priority_row = child.xpath(priority_query, namespaces = doc_ns)
            priority_cells = priority_row[0].xpath(".//w:tc",
                                                   namespaces = doc_ns)
            priority_para = priority_cells[1].xpath(".//w:t",
                                                    namespaces = doc_ns)
            priority = "".join([string.text for string in priority_para])

            # We'll add one more value to this dict later, when we extract
            # the report date from the footer.
            details = {"ticket": ticket, "site": site, "priority": priority}

            # Finally, reset the "details_found" flag to false, so that we
            # don't try to extract data from subsequent tables.
            details_found = False

        # Check for the "Details" title. Unless the "Details" title is
        # found, this will be an empty list, which will evaluate as False
        # in the Boolean test above.
        details_found = child.xpath('.//w:t[text()="' + DETAILS_TITLE + '"]',
                                    namespaces = doc_ns)

        # If the previous child was the "Updates" title, search the table
        # for updates, and then break the loop.

        if updates_found:

            # Find all of the update cells.
            updates_query = './/w:tc[descendant::*[contains(text(), "' + \
                            STATUS_CHANGE_PHRASE + '")]]'
            updates = child.xpath(updates_query, namespaces = doc_ns)

            for update in updates:

                # Get all the text paragarphs in the update cell.
                update_paras = update.findall("w:p", namespaces = doc_ns)

                # The status change is always the first paragraph in an
                # update, and follows a set format. Given that the time
                # paragraph is sometimes split across multiple text
                # elements (see below), it's possible the same is true of
                # the status change paragraph, and so we deal with this
                # possibility.
                status_para = update_paras[0].findall(".//w:t",
                                                      namespaces = doc_ns)
                status_para_string = "".join([string.text for string
                                              in status_para])

                # Get the new status.
                to_status_split = status_para_string.split(TO_SEP)
                to_status = to_status_split[-1]

                # In some cases, the new status (but not the old one)
                # includes a parenthetical note. If this is present,
                # split it out.
                status_note_split = status_note_patt.split(to_status)
                if len(status_note_split) > 1:
                    to_status = status_note_split[0]
                    status_note = status_note_split[1]
                else:
                    status_note = "None"

                # Get the old status, if any.
                from_status_split = to_status_split[-2].split(FROM_SEP)
                if len(from_status_split) > 1:
                    from_status = from_status_split[-1]
                else:
                    from_status = "None"

                # Get the updater.
                updater_split = from_status_split[0].split(UPDATER_SEP)
                updater = updater_split[0]

                # The time is always the last paragraph in an update
                # cell. However, this one-line paragraph is sometimes
                # split across two (or possibly more) text elements.
                time_para = update_paras[-1].findall(".//w:t",
                                                     namespaces = doc_ns)
                time_string = "".join([string.text for string in time_para])

                # Pull out the entry time.  in some cases, this is also
                # the effective time (in which case the split returns a
                # single-element list with nothing but the entry time),
                # while in other cases, the two are listed separately.
                time_string_split = time_string.split(EFFECTIVE_TIME_SEP)
                entry_time_string = time_string_split[0]
                entry_time = datetime.datetime.strptime(entry_time_string,
                                                        TIME_FORMAT)

                # Pull out the effective time. We pull out the last sub-
                # (see the above comment) string, and strip a trailing
                # parenthesis if it's present.
                eff_time_string = \
                    time_string_split[-1].rstrip(EFFECTIVE_TIME_STRIP_CHARS)
                eff_time = datetime.datetime.strptime(eff_time_string,
                                                      TIME_FORMAT)

                # Add the extracted values to the updates list.
                d = {"updater": updater, "entry time": entry_time,
                     "from status": from_status, "to status": to_status,
                     "status note": status_note, "effective time": eff_time}
                updates_list.append(d)

            # If we've found and searched the "Updates" table, we're done.
            break

        # Check for the "Updates" title. Unless the "Updates" title is
        # found, this will be an empty list, which will evaluate as False
        # in the Boolean test above.
        updates_found = child.xpath('.//w:t[text()="' + UPDATES_TITLE + '"]',
                                    namespaces = doc_ns)

    # Reverse the list--in most cases, this should give us the proper
    # ordering.
    updates_list.reverse()

    # Just in case, sort the list by effective time--notably, this won't
    # change the ordering of updates with the same effective time (which
    # will prevent two updates entered in quick succession from being
    # reversed). If effective times are correct, this should give the
    # proper ordering almost 100% of the time, with the sole exception
    # being status changes with the same time that are entered out of
    # order.
    updates_list.sort(key = lambda update: update["effective time"])

    # Now, we need to calculate delays between updates, and between
    # effective and report times.

    old_update = None
    for update in updates_list:

        # If this is the first update, the delay between updates is defined
        # as 0.
        if not old_update:
            update["status hours"] = 0

        # Otherwise, calculate the delay--unless the old status indicated
        # the clock was suspended.
        #
        # We check here to make sure the "to status" of the old update and
        # the "from status" of the current update are identical--they
        # should be, but if they aren't (see the above comment on the list
        # sort), we record an error code in place of the delay.
        #
        # Since "updates_list" is actually a list of pointers to individual
        # dicts, we can modify each of those individual dicts ("update") in
        # turn without causing problems.
        else:
            if update["from status"] != old_update["to status"]:
                update["status hours"] = "Status mismatch!"
            elif update["from status"] in SUSPEND_STATUSES:
                update["status hours"] = 0
            else:
                update["status hours"] = \
                    calculate_delay(old_update["effective time"],
                                    update["effective time"],
                                    priority)

        # Calculate the delay between the entry time and the effective
        # time.
        update["update delay"] = calculate_delay(update["effective time"],
                                                 update["entry time"],
                                                 priority)

        # Replace "old_update" with the current update.
        old_update = update

    # Find the report time in the footers, and add it to the details list.
    # As with the other strings, we account for the possibility of the
    # string being spread across multiple text elements.
    report_time_query = './/w:p[descendant::*[starts-with(text(), "' + \
                        REPORT_TIME_PREFIX + '")]]'
    report_time_para = footer.xpath(report_time_query,
                                    namespaces = footer_ns)[0]
    report_time_para_text = report_time_para.findall(".//w:t",
                                                     namespaces = doc_ns)
    report_time_entry_string = "".join([string.text
                                        for string in report_time_para_text])
    report_time_string = report_time_entry_string.split(REPORT_TIME_PREFIX)[-1]
    report_time = datetime.datetime.strptime(report_time_string, TIME_FORMAT)
    details["report time"] = report_time

    return details, updates_list


# Get the filename from the command line. This could be either a DOCX file
# (for a single ticket) or a a ZIP of a DOCX files (for multiple tickets).
try:
    filename = argv[1]
except IndexError:
    print("ERROR: No input file specified.")
    stdout.flush()
    sleep(5)
    sys.exit()

# Read the file--it's going to be in ZIP format either way.
try:
    input_file = ZipFile(filename)
except Exception as e:
    print("Error opening Word document: " + filename)
    print(e)
    stdout.flush()
    sleep(5)
    sys.exit()

# To determine what type of file we're dealing with, we'll use the
# file extension.
file_ext = splitext(filename)[-1].lstrip(".").lower()

# A single ticket will be a DOCX file.
if file_ext == "docx":

    # Get the updates data and details from the ticket.
    details, updates_list = parse_ticket(input_file)

    # Add a line to "updates_list" that shows the delay between the
    # completion time and the report time.
    report_time = details["report time"]
    report_delay = calculate_delay(updates_list[-1]["effective time"],
                                   report_time, details["priority"])
    report_time_update = {"updater": "", "entry time": report_time,
                          "from status": updates_list[-1]["to status"],
                          "to status": "Report Date", "status note": "None",
                          "effective time": report_time,
                          "status hours": report_delay, "update delay": ""}
    updates_list.append(report_time_update)

    # Write the output CSV.
    report_filename = filename.replace('docx','csv')
    write_single_ticket_CSV(report_filename, details, updates_list)

# A collection of tickets will be a ZIP file containing a directory with a
# single DOCX file for each ticket.
elif file_ext == "zip":

    # Initialize lists of details and updates.
    details_list = []
    all_updates = []

    # Iterate through the individual ticket files, getting the details and
    # updates for each; we add the ticket number to each of the updates, as
    # a key to cross-refernece the two lists.
    #
    # The "if" skips over directories (The list of directories and files in
    # a "ZipFile" object is flat, not hierarchical.)
    for zipped_file in input_file.filelist:
        zipped_file_name = zipped_file.filename
        if splitext(zipped_file_name)[-1].lstrip('.').lower() == "docx":
            word_file = ZipFile(BytesIO(input_file.read(zipped_file_name)))
            details, updates_list = parse_ticket(word_file)
            details_list.append(details)
            ticket = details["ticket"]
            for update in updates_list:
                update["ticket"] = ticket
            all_updates.extend(updates_list)

    # Write the details and updates CSV's.
    details_filename = filename.replace('.zip','_details.csv')
    updates_filename = filename.replace('.zip','_updates.csv')
    write_details_CSV(details_filename, details_list)
    write_updates_CSV(updates_filename, all_updates )

else:
    print('File type "' + file_ext + '" is not supported.')
    stdout.flush()
    sleep(5)
    sys.exit()

print("\nRun completed.\n")
