import requests

data ={'status': [],
       'country': [],
       'countryCode': [],
       'region': [],
       'regionName': [],
       'city': [],
       'zip': [],
       'lat': [],
       'lon': [],
       'timezone': [],
       'provider': [],
       'organization': [],
       'as': [],
       'query IP': []}

def get_info_by_ip(ip="127.0.0.1"):
    try:
        response = requests.get(url = f"http://ip-api.com/json/{ip}").json()
        data['status'].append(response.get('status'))
        data['country'].append(response.get('country'))
        data['countryCode'].append(response.get('countryCode'))
        data['region'].append(response.get('region'))
        data['regionName'].append(response.get('regionName'))
        data['city'].append(response.get('city'))
        data['zip'].append(response.get('zip'))
        data['lat'].append(response.get('lat'))
        data['lon'].append(response.get('lon'))
        data['timezone'].append(response.get('timezone'))
        data['provider'].append(response.get('isp'))
        data['organization'].append(response.get('org'))
        data['as'].append(response.get('as'))
        data['query IP'].append(response.get('query'))
    except:
        print("!! Please check your connection !!")

def main():
    ip = "212.30.134.176"
    get_info_by_ip(ip)
    print(data)

if __name__ == "__main__":
    main()