import requests
import os
import datetime
import gzip
import shutil
import numpy as np
import pandas as pd
import math
import json


class scrapnvd:
    def __init__(self,year,start_year,end_year):
        self.year = year
        self.start_year = start_year
        self.end_year = end_year

    def extract_gzip(self,input_file, output_file):
        with gzip.open(input_file, 'rb') as f_in:
            with open(output_file, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)

    def download_nvd_data(self):
        nvd_url = f"https://nvd.nist.gov/feeds/json/cve/1.1/nvdcve-1.1-{self.year}.json.gz"
        file_name = f"nvd_data_{self.year}.json.gz"
        response = requests.get(nvd_url)
        if response.status_code == 200:
            with open(file_name, 'wb') as file:
                file.write(response.content)
            print(f"\nDownloaded NVD data for {self.year} to {file_name}")
            print(f"Extracting Downloaded File to {self.year}.json")
            output_file = f"{self.year}.json"
            self.extract_gzip(file_name,output_file)
            print("File Extracted Successfully")
        else:
            print(f"\nFailed to download NVD data. Status code: {response.status_code}")

    def create_excel_nvd_data(self):
        print(f"Creating Excel of year {self.year}")
        file = f'{self.year}.json'
        data = pd.read_json(file, encoding='utf-8') 
        updated_cve_data = pd.json_normalize(data['CVE_Items'])
        updated_cve_data.to_excel(f"{self.year}.xlsx",sheet_name='CVE-DATABASE',index=False)
        print("Created excel NVD Data Successfully")

    def combine_data(self):
        print("\nCreating Combined NVD Database...")
        final_cve_data = pd.DataFrame()
        for i in range(self.start_year, self.end_year+1):
            data = pd.read_excel(f"{i}.xlsx")
            final_cve_data = pd.concat([final_cve_data,data], axis=0)
        final_cve_data.to_excel("NVD-CVE-DATABASE.xlsx",sheet_name='CVE-DATABASE',index=False)
        print("\nFile created successfully. Please check NVD-CVE-DATABASE.xlsx")

if __name__ == "__main__":
    start_year = 2005
    end_year = 2024
    for i in range(start_year, end_year+1):
        obj = scrapnvd(i,start_year,end_year)
        obj.download_nvd_data()
        obj.create_excel_nvd_data()

    obj = scrapnvd("2000",start_year,end_year)
    obj.combine_data()