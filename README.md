Ansible module which imports an Excel spreadsheet to Ansible facts. 

Each sheet is represented by "spreadsheet_{sheet_name}" in Ansible.  The column headers of each row in the sheet are converted to variable names,  and values are populated for each row.  

Example:

"spreadsheet_Access": [
        {
            "Access_VLAN": null,
            "Description": "To 5600-1",
            "Hostname": "Leaf-1",
            "Interface": "Ethernet1/6",
            "Intf_Type": "trunk",
            "is_vPC_Member": "True",
            "vPC_Nbr": 10
        },
        {
            "Access_VLAN": null,
            "Description": "To 5600-1",
            "Hostname": "Leaf-2",
            "Interface": "Ethernet1/6",
            "Intf_Type": "trunk",
            "is_vPC_Member": "True",
            "vPC_Nbr": 10
        },
        {
            "Access_VLAN": null,
            "Description": "To 5600-2",
            "Hostname": "Leaf-3",
            "Interface": "Ethernet1/1",
            "Intf_Type": "trunk",
            "is_vPC_Member": "False",
            "vPC_Nbr": null
        }
    ]
}

 
