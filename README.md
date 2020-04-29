# Country_Channel_Mappings
This simple script is designed to pull all of the channels allowed by country for Mist APs.

##Requirments:
Python 3.5+
### Libraries:
- requests
- xlsxwriter
### Informaiton:
There are 3 requirements for the script to run.
- Mist API Token
- Mist Org ID
- Mist Site ID

These are passed in as CLI parameters

## Usage:
```bash
python country_mappings.py -k <api key> -o <org_id> -s <site_id>
```
## Help info
For additional options or info, you can run:
```bash
python country_happings.py --help
```