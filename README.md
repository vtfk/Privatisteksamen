# Privatisteksamen

## config.json

Create a config.json file with the following structure. Swap out the values with your own

```json
{
    "path": "folder-path-where-your-lists-live",
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

Filenames must include the date for when the privatist exam is held. The filename can include text before and/or after the date (if someone wants that).

File types must be one of:
* .csv - Must use ';' as separator
* .xlsx

Examples:
* "Norsk skriftlig 09.11.2020 Blått rom.csv"
* "09.11.2020_Norsk_Rødt_Rom.xlsx"
* "09.11.2020.csv"

## Example files

Look at the example files in the *examples* folder. Use them as a starting point

## **Caution:**

**Excel automatically removes the first 0 from personal numbers which starts with a 0!**

Don't worry. The script takes this into account! :sunglasses: