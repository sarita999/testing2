import requests
from shareplum import Site
from requests_ntlm import HttpNtlmAuth
 
user_credentials = {
    'username' : 'sharmila.konreddy',#Application host id
    'password' : 'a.konr@123',#application host id password
    'domain' : 'myconcerto'
}
 
# Creating a class for Authentication
class UserAuthentication:
 
    def __init__(self, username, password, domain, site_url):
        self.__username = username
        self.__password = password
        self.__domain = domain
        self.__site_url = site_url
        self.__ntlm_auth = None
 
    def authenticate(self):
        login_user = self.__domain + "\\" + self.__username  # username example: myconcerto\sharmila.konreddy
        user_auth = HttpNtlmAuth(login_user, self.__password)
        self.__ntlm_auth = user_auth
 
        # Create header for the http request
        my_headers = {
            'accept' : 'application/json;odata=verbose',
            'content-type' : 'application/json;odata=verbose',
            'odata' : 'verbose',
            'X-RequestForceAuthentication' : 'true'
            }       
         
        # Sending http get request to the sharepoint site and Requests ignore verifying the SSL certificates if you set verify to False
        result = requests.get(self.__site_url, auth=user_auth, headers=my_headers, verify=False)
        
        # Checking the status code of the requests
        if result.status_code == requests.codes.ok:  # Value of requests.codes.ok is 200
            site = Site(self.__site_url, auth=user_auth)
            sp_list = site.List('Scrubbing Requests')
            data = sp_list.GetListItems('All Items')            
            return data
        else:
            result.raise_for_status()
 
if __name__ == "__main__":
    username = user_credentials['username']
    password = user_credentials['password']
    domain = user_credentials['domain']
    #site_url = "http://azr0004wspws1/sites/DigitalAssetManagement/_api/web/lists/GetByTitle('Scrubbing Request')"
    site_url="https://stg.myconcerto.accenture.com/sites/digitalassetmanagement"
    auth_object = UserAuthentication(username, password, domain, site_url)
    result = auth_object.authenticate()
    if result:
        print(result)
        print("Successfully login to sharepoint site")
