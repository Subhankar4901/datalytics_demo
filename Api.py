import sys
from jmespath import search
import requests
import json
import datetime
from decouple import config
class BaseAPI:
    _headers={
        "x-rapidapi-key":config("API_KEY"),
        "x-rapidapi-host":config("HOST")
    }
    
    @classmethod
    def base_fetch(cls,query:dict,url:str=None):
        response=requests.get(url=url,params=query,headers=cls._headers)
        try:
            return response.json()
        except:
            print(response.text)
            print(response.content)
            sys.exit(0)
    @classmethod
    def base_output(cls,query:dict,property:str=None):
        # loading config from api_config
        with open("api_config.json","r") as f:
            api_config=json.load(f)
        keys=api_config[property]["keys"]
        url=api_config[property]["url"]
        # fetching raw data from api
        raw_data:dict=cls.base_fetch(query=query,url=url)
        data={}
        # structuring and filtering the data
        for year_wise_property in raw_data[keys[0]][keys[1]]:
            year=datetime.datetime.strptime(year_wise_property["endDate"]["fmt"],"%Y-%m-%d").year
            data[str(year)]={}
            for financial_property in year_wise_property:
                if financial_property=="endDate" or financial_property=="maxAge":
                    continue
                try:
                    data[str(year)][financial_property]=year_wise_property[financial_property]["fmt"]
                    # There are blank financial atributes in api output
                except:
                    continue
        return data
    @classmethod
    def search(cls,data,required_key):
        if required_key in data:
            return data[required_key]
        for key in data:
            if isinstance(data[key],dict):
                var=cls.search(data[key],required_key)
                if var is not None:
                    return var
        return None
        
        