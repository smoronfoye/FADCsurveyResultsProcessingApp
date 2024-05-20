import pandas as pd
import requests

# from glob import glob

MAX_LRU_SIZE = 100
base_url = "https://maps.googleapis.com/maps/api/geocode/json?latlng="

lru_cache_map = dict()
lru_cache_array = []


def processExcelFile(fileName, sheetName, latitudeColumn, longitudeColumn):
    df = pd.read_excel(fileName, sheet_name=sheetName)
    df = df.sort_values([latitudeColumn, longitudeColumn], ascending=False)

    file_name_without_type = fileName.split(".")[0]

    rows_list = []
    for index, row in df.iterrows():

        latitude = row[latitudeColumn]
        longitude = row[longitudeColumn]

        api_url = base_url
        api_url += str(latitude) + "," + str(longitude)
        api_url += "&key=AIzaSyBcFFE4HYe5up8W6TBMnZe4IG2FINWMf_A"

        response = {}
        if latitude and longitude:
            key = (latitude, longitude)
            response_from_cache = readFromLruCache(key)
            if response_from_cache:
                response = response_from_cache
            else:
                resultTypes = [
                    "street_address",
                    "premise",
                    "establishment",
                    "point_of_interest",
                    "neighborhood",
                ]
                for resultType in resultTypes:
                    if len(response) == 0 or len(response["results"]) == 0:
                        response = requests.get(api_url + "&result_type=" + resultType)
                        response = response.json()
                    else:
                        break

                updateLruCache(key, response)

        street_address = ""
        place_id = ""
        if len(response) == 0 or len(response["results"]) > 0:
            street_address = response["results"][0]["formatted_address"]
            place_id = response["results"][0]["place_id"]

        row["street_address"] = street_address
        row["place_id"] = place_id

        rows_list.append(row)

    processedSurveyResults = pd.DataFrame(rows_list)
    processedSurveyResults.to_csv(
        file_name_without_type + "_processed.csv", sep=",", encoding="utf-8"
    )


# def main():
#     for file_name in glob('survey_results_*.xlsx'):
#         processExcelFile(file_name, 'Sheet1', 'Cordinate-Latitude', 'Cordinate-Longitude')


def readFromLruCache(key: tuple):
    if key in lru_cache_map:
        lru_array_index = lru_cache_map[key]
        return lru_cache_array[lru_array_index]
    else:
        return None


def updateLruCache(key: tuple, value):
    # stubed becuase we definietely need a linked list to implement
    # if key in lru_cache_map:
    #     lru_array_index = lru_cache_map[key]
    #     del lru_cache_array[lru_array_index:lru_array_index]

    # if len(lru_cache_array) >= MAX_LRU_SIZE:
    #     del lru_cache_array[0]
    #     new_first_element = lru_cache_array[0]
    #     my_dict['key']
    if (
        len(lru_cache_array) < MAX_LRU_SIZE
    ):  # this wll need to be removed if eventually implmneted with a linked list
        index = len(lru_cache_array)
        lru_cache_array.append(value)
        lru_cache_map[key] = index


# if __name__ == "__main__":
#     main()


# pyinstaller --onefile --clean --upx-dir C:\Users\kmoro\Downloads\upx-4.2.2-win64\upx-4.2.2-win64 --icon favicon.ico  --hidden-import openpyxl.cell._writer user-interface.py


# pyinstaller --onefile --clean --upx-dir /home/seyi/upx-4.2.2-amd64_linux --icon favicon.ico  --hidden-import openpyxl.cell._writer user-interface.py
