#!/usr/bin/env python3

import requests
import json
import sys
import xlsxwriter


with open("channel_mappings.json", 'r') as channel_map:
    channel_mapping = json.load(channel_map)


class MistAPI(object):
    def __init__(self, host: str, org: str):
        self.host = host
        self.org = org
        self.header = ""


class MistAPIToken(MistAPI):
    def __init__(self, host: str, org: str, mist_api_token: str):
        """

        :param host: api host: ex api.mist.com
        :param org: org_id: example xxxxxxx-xxxx-xxx-xxxxxxxxx
        :param mist_api_token: mist API token "xxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
        """
        super(MistAPIToken, self).__init__(host, org)
        self.mist_api_token = mist_api_token
        self.header = {"Authorization": f"Token {mist_api_token}"}


class Mist:
    def __init__(self, mist_api: MistAPI):
        """

        :param mist_api: MistAPI object or inherited method like MistAPIToken
        """
        self.mistAPI = mist_api
        if not self.verify_self:
            raise ValueError("Please verify mist authentication and try again")


    def http_get(self, url):
        """

        :param url: url extension.  Example: /api/v1/self
        :return: requests response object
        """
        try:
            header = {**{"content-type": "application/json"}, **self.mistAPI.header}
            my_url = f"https://{self.mistAPI.host}{url}"
            response = requests.get(my_url, headers=header)
            return response
        except Exception as e:
            print(e)
            return None

    def http_post(self, url: str, body: dict):
        """

        :param url: url extension.  Example: /api/v1/self
        :param body: dictionary formatted body for a post
        :return:
        """
        response = None
        try:
            header = {**{"content-type": "application/json"}, **self.mistAPI.header}
            my_url = f"https://{self.mistAPI.host}{url}"
            response = requests.post(my_url, headers=header, data=json.dumps(body))
        except Exception as e:
            print(e)
        return response

    def verify_self(self):
        """

        :return: verifies that API credential successfully return a /api/v1/self
        """
        try:
            results = self.http_get("/api/v1/self")
            if results.status_code == 200:
                return True
        except Exception as e:
            print(e)
            return False

    def get_rf_templates(self):
        """

        :return: returns a list of dictionay rf_templates
        """
        rf_templates = None
        try:
            rf_templates = self.http_get(f"/api/v1/orgs/{self.mistAPI.org}/rftemplates").json()
        except Exception as e:
            print(e)
        return rf_templates

    def get_rftemplate_by_name(self, rf_template_name: str):
        """

        :param rf_template_name: Name of the RF Template
        :return: rf template dictionary
        """
        rf_templates = self.get_rf_templates()
        return next(item for item in rf_templates if item["name"] == rf_template_name)

    def create_site(self, body):
        """

        :param body: properly formatted body for a site creation
        :return: requests response object
        """
        response = self.http_post(f"/api/v1/orgs/{self.mistAPI.org}/sites", body)
        return response


def get_parser():
    """
    :return: parser for argparse
    """
    from argparse import ArgumentParser
    parser = ArgumentParser(description="Mist site creation tool")
    parser.add_argument(
        "-k", "--key", dest="mist_api_key", help="Mist API Key", type=str, required=True
    )
    parser.add_argument(
        "-o", "--org", dest="org_id", help="Mist Org ID", type=str, required=True
    )
    parser.add_argument(
        "-e", "--EU", dest="mist_europe", help="Mist EU Environment", required=False
    )
    parser.add_argument(
        "-s", "--site", dest="site_id", help="Site ID for checking Country info", required=True
    )
    return parser


def get_channel_column(channel):
    """

    :param channel: wifi channel number as integer
    :return: excel column letter
    """
    my_channel = str(channel)
    global channel_mapping
    return channel_mapping[my_channel]

def main(argv):

    try:
        org_id = argv.org_id
        mist_api_key = argv.mist_api_key
        site_id = argv.site_id
    except:
        print("Missing Required Values, aborting")
        sys.exit()
    mist_api = MistAPIToken("api.mist.com", org_id, mist_api_key)
    try:
        mist_connector = Mist(mist_api)
    except ValueError:
        sys.exit()
    country_const = mist_connector.http_get("/api/v1/const/countries").json()

    results = []
    for entry in country_const:
        results.append(mist_connector.http_get(f"/api/v1/sites/{site_id}/devices/ap_channels?country_code={entry['alpha2']}").json())

    build_xlsx(results)


def build_xlsx(results, file_name="country_channels.xlsx"):
    """

    :param results: list of outputs from /api/v1/mist/site/:site_id/devices/ap_channels?country_code={alpha2}
    :param file_name: Name of the excel file to be created
    """
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    worksheet.write("A1", "Name")
    worksheet.write("B1", "Alpha2")
    worksheet.write("C1", "DFS_OK")
    worksheet.write("D1", "Band24_Enabled")
    worksheet.write("E1", "Band24_40mhz_Allowed")
    worksheet.write("F1", "Band5_Enabled")
    worksheet.write("G1", "Certified")
    worksheet.write("H1", "Uses")
    for key in channel_mapping.keys():
        worksheet.write(f'{get_channel_column(key)}1', key)
    for entry in results:
        row = results.index(entry) + 2
        worksheet.write(f"A{row}", str(entry['name']))
        worksheet.write(f"B{row}", str(entry['key']))
        if "dfs_ok" in entry.keys():
            worksheet.write(f"C{row}", str(entry['dfs_ok']))
        else:
            worksheet.write(f"C{row}", 'False')
        worksheet.write(f"D{row}", str(entry['band24_enabled']))
        worksheet.write(f"E{row}", str(entry['band24_40mhz_allowed']))
        worksheet.write(f"F{row}", str(entry['band5_enabled']))
        if "certified" in entry.keys():
            worksheet.write(f"G{row}", str(entry['certified']))
        else:
            worksheet.write(f"G{row}", "False")
        if "uses" in entry.keys():
            worksheet.write(f"H{row}", str(entry['uses']))
        if entry['band24_enabled']:
            for channel in entry['band24_channels']['20']:
                worksheet.write(f'{get_channel_column(channel)}{row}', 'X')
        if entry['band5_enabled']:
            for channel in entry['band5_channels']['20']:
                worksheet.write(f'{get_channel_column(channel)}{row}', 'X')
    workbook.close()
    return


if __name__ == '__main__':
    my_parser = get_parser()
    my_args = my_parser.parse_args()
    main(my_args)