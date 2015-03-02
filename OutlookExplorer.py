#!/usr/bin/env python

"""
A simple address book browser. Descend into the depths of
your Exchange Server through Outlook and find that contacts
list. Then extract those contacts!

NOTE: Generate the type library for Microsoft Outlook first.

Mail Messages
-------------

In MSDN, the interfaces are documented under "Messaging and
Collaboration", "Collaboration Data Objects", "CDO for
Windows 2000", "CDO for Windows 2000", "Messaging",
"Concepts", "CDO Messaging COM Interfaces", "The IMessage
COM Interface", "IMessage Properties".

Unfortunately, not all properties are always available.
Indeed, the 'gen_py' directory inside 'site-packages' may be
a more accurate summary of properties and interfaces.

Saved Items
-----------

    typedef enum {
        olTXT = 0,
        olRTF = 1,
        olTemplate = 2,
        olMSG = 3,
        olDoc = 4,
        olHTML = 5,
        olVCard = 6,
        olVCal = 7
    } OlSaveAsType;

Appointment items are saved as vCal by default.
Contact items are saved as vCard by default.
Mail items are saved as text by default.
All other items are saved as text by default.
"""

import win32com.client
import sys, os

# This is needed before we can even talk about different enum values.

outlook = win32com.client.Dispatch("Outlook.Application")

class View:

    "A view onto namespaces."

    def __init__(self, encoding):

        "Initialise the view with a character 'encoding'."

        self.encoding = encoding

class ConsoleView(View):

    "A console-style view."

    show_namespace_property_mapping = {
        win32com.client.constants.olFolder :
            ("+", "Name"),
        win32com.client.constants.olContact :
            (">", "Email1Address"),
        win32com.client.constants.olMail :
            (">", "Subject"),
        None :
            ("?", "Name")
        }

    def __init__(self, encoding, page_width=80, page_limit=20):

        """
        Initialise the view with a character 'encoding' and the optional
        'page_width' and 'page_limit'.
        """

        View.__init__(self, encoding)
        self.page_width = page_width
        self.page_limit = page_limit

    def update_status(self, counter, max_value):

        "Update a status indicator with the given 'counter' value and 'max_value'."

        last_counter = max(counter - 1, 0)
        last_width = int((last_counter * self.page_width) / max_value)
        width = int((counter * self.page_width) / max_value)

        sys.stdout.write("." * (width - last_width))

        if counter == max_value:
            sys.stdout.write("\n")

    def error(self):
        sys.stdout.write("!")

    def show_namespace(self, items):

        "Show the namespace, given a list of 'items'."

        if len(items) > self.page_limit:
            print "!", "Showing the first", self.page_limit, "items only."

        for value in items[:self.page_limit]:
            try:
                decoration, property = self.show_namespace_property_mapping[value.Class]
            except KeyError:
                decoration, property = self.show_namespace_property_mapping[None]

            print decoration, self.get_property(value, property).encode(self.encoding)

    def get_property(self, object, property, default=""):
        try:
            # NOTE: Hack!

            if property == "SentOn":
                return getattr(object, property).Format()

            return getattr(object, property)

        except AttributeError:
            return default

class Extractor:

    "A class used for the extraction of items/objects from folders."

    extract_type_mapping = {
        win32com.client.constants.olAppointment :
            (win32com.client.constants.olVCal, "vcs"),
        win32com.client.constants.olContact :
            (win32com.client.constants.olVCard, "vcf"),
        win32com.client.constants.olMail :
            (win32com.client.constants.olTXT, "txt"),
        None :
            (win32com.client.constants.olTXT, "txt")
        }

    def __init__(self, view=None):

        "Initialise the extractor with the optional 'view'."

        self.view = view

    def extract(self, items, filename):

        "Extract the given 'items' to a file with the given 'filename'."

        total_number = len(items)
        for i in range(0, total_number):
            value = items[i]

            try:
                save_as_type, suffix = self.extract_type_mapping[value.Class]
            except KeyError:
                save_as_type, suffix = self.extract_type_mapping[None]

            try:
                value.SaveAs(os.path.join(filename, str(i) + "." + suffix),
                    save_as_type)
            except AttributeError:
                if self.view:
                    self.view.error()
            except win32com.client.pywintypes.com_error:
                if self.view:
                    self.view.error()

            if self.view:
                self.view.update_status(i + 1, total_number)

class Explorer:

    "A class maintaining the state of exploration."

    def __init__(self, view=None):
        global outlook
        self.current = self.ns = outlook.GetNamespace("MAPI")
        self.view = view
        self._get_namespace()

    def up(self):

        "Ascend into the parent folder returning whether it was possible."

        if self.current != self.ns:
            self.current = self.current.Parent
            self._get_namespace()
            return 1
        return 0

    def down(self, name):

        """
        Descend into the folder with the given 'name' returning whether it
        could be done.
        """

        if self.choices.has_key(name):
            self.current = self.choices[name]
            self._get_namespace()
            return 1
        return 0

    def get_items(self):

        "Return a list of items in the current folder."

        return self.items

    def get_choices(self):

        "Return a dictionary mapping names to objects."
        
        return self.choices

    def _get_namespace(self):

        """
        An internal method which refreshes the current namespace.
        """
        
        self.choices, self.items = get_namespace(self.current, self.view)

def get_namespace(namespace, view=None):

    """
    Get the contents of the given 'namespace', returning a dictionary of
    choices (appropriate for folders) and a list of items (appropriate for
    messages).
    """

    d = {}
    l = []

    # First try looking for folders. Then look for items. And so on.

    for properties in (("Folders", "Name"), ("Items", None)):

        # Look for objects of the current type: folders, items, etc.

        object_name = properties[0]

        try:
            subobject = getattr(namespace, object_name)
        except AttributeError:
            # Ignore the rest of the loop body and start
            # the next iteration.
            continue

        # Index the retrieved items by storing them by name in a dictionary.
        # Cannot slice items, and they always seem to start at index 1.

        total_number = len(subobject)
        for i in range(1, total_number + 1):
            try:
                field_name = properties[1]

                # Store in the dictionary using the specified property, if
                # specified.

                l.append(subobject[i])
                if field_name is not None:
                    d[getattr(subobject[i], field_name)] = subobject[i]

            except AttributeError:
                pass

            # Crude status indicator.

            if view:
                view.update_status(i, total_number)

    return d, l

def main():

    # Get the encoding if specified.

    if len(sys.argv) > 2:
        encoding = sys.argv[2]
    else:
        encoding = "UTF-8"

    view = ConsoleView(encoding)
    explorer = Explorer(view)

    while 1:
        # Prompt the user.
        print "-" * 60
        view.show_namespace(explorer.get_items())
        print "-" * 60
        print "[U]p [D]own [E]xtract [Q]uit [H]elp then <Return>!"
        print "-" * 60
        s = raw_input().strip().upper()[0]

        # Find the right action.        
        if s == "U":
            # Up!
            explorer.up()

        elif s == "D":
            # Prompt for the folder to enter.
            print "Down into:"
            name = raw_input()
            if not explorer.down(name):
                print "No such object."
                
        elif s == "E":
            # Prompt for the file to extract to.
            print "Extract to:"
            filename = raw_input()
            print "Extracting..."
            extractor = Extractor(view)
            extractor.extract(explorer.get_items(), filename)
            
        elif s == "Q":
            print "Exiting..."
            raise SystemExit

        elif s == "H":
            print "Type the D key then <Return> to enter a folder."
            print "Give the exact folder name when prompted and press <Return>."
            print "Type the U key then <Return> to move up one level."
            print "Type the Q key then <Return> to quit."
            print "Type the H key then <Return> to read this again."
            print "Type the E key then <Return> to extract items."
            print "Give a filename when prompted and press <Return>."
            print
            print "Good luck!"
        else:
            print "No such command."

if __name__ == "__main__":
    main()

# vim: tabstop=4 expandtab shiftwidth=4
