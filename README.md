# Privatisteksamen

## config.json

Create a config.json file with the following structure. Swap out the values with your own

```json
{
    "path": "folder-path-where-your-lists-live",
    "includedExtensions": [
        "*.xlsx",
        "*.csv",
        "*.xls"
    ],
    "excludedFolders": [
        "Finished",
        "Test",
        "Backup"
    ],
    "finishedFolder": "Finished",
    "orgName": "name-of-organization",
    "ad": {
        "searchBase": "distinguishedName-of-the-ou-to-search-for-users",
        "group": "name-of-the-group-to-add/remove-users-to/from",
        "server": "hostname/ipaddress-of-AD-controller",
        "enabledUsersOnly": false
    },
    "smtp": {
        "server": "hostname/ipaddress-of-your-smtp-server",
        "fromAddress": "privatistinfo@vtfk.no",
        "fromDisplayName": "Privatistinfo",
        "to": [
            {
                "address": "jule.duk@vtfk.no",
                "displayName": "Jule Duk"
            },
            {
                "address": "ape.loff@vtfk.no",
                "displayName": "Ape Loff"
            }
        ],
        "cc": [
            {
                "address": "stor.kar@vtfk.no",
                "displayName": "Stor Kar"
            }
        ],
        "bcc": [
            {
                "address": "jule.nissen@vtfk.no",
                "displayName": "Jule Nissen"
            }
        ]
    }
}
```

## Filename

Filename(s) can be whatever.

File types must be one of:
* .csv - `Must use ';' as separator`
* .xlsx
* .xls - `Caution! This file type will automatically be converted to .xlsx and the converted file will be used!`

## Content

File(s) must have these headers:

* Eksamensparti
* Eksamensdato
* Fødselsnummer

## Example files

Look at the example files in the *examples* folder. Use them as a starting point

## **Caution:**

**Excel automatically removes the first 0 from personal numbers which starts with a 0!**

Don't worry. The script takes this into account! :sunglasses:

## End of exam period

When the exam period is over (all `Eksamensdato` entries are in the past), the exam file will automatically be moved to `finishedFolder` set in config file.

If there's no more files/folders in the folder where the exam file was located, the folder will automatically be removed!
