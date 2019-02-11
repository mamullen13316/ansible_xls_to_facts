Ansible module which imports an Excel spreadsheet to Ansible facts.

Each sheet is represented by "spreadsheet_{sheet_name}" in Ansible.  The column headers of each row in the sheet are converted to variable names,  and values are populated for each row.

Copy to the library path as specified in your ansible.cfg file.

Requires the openpyxl Python module on the Ansible host.
 - Default behavior is set for [*data_only*](https://openpyxl.readthedocs.io/en/stable/changes.html#id485)  to allow extracting values only from formulae

Install with PIP:

sudo pip install openpyxl
