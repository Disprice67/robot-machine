from bs4 import BeautifulSoup
from ebaysdk.finding import Connection as Finding
from ebaysdk.trading import Connection as Trading
from requests_html import HTMLSession

from config import URL_HUAWEI, USER_AGENT


class Ebay:

    """
    This class works with ebay and website positions
    -----
    * method searchebay() allows you to work with ebay through the API,
    and pulls the necessary information on positions on the site
    """

    def __init__(self, api: str, cert_id: str,
                 dev_id: str, token: str):

        self.api_key = api
        self.cert_id = cert_id
        self.dev_id = dev_id
        self.token = token

    def searchebay(self, key: str, data: dict, 
                   excel: 'Excel', vendor: str, main_key: str):
        """Search_Ebay."""

        data[main_key]['URL'] = 'Нет результатов.'
        data[main_key]['СТОИМОСТЬ ТОВАРА/USD'] = 0
        
        if vendor != 'None':
            filter_vendor = excel.filterkey(vendor, 'PN')
            search = vendor + ' ' + key
        else:
            search = key

        payload: dict = {'keywords': f'{search}',
                         'paginationInput': {'entriesPerPage':15},
                         'itemFilter': [{'name': 'LocatedIn',
                                         'value': 'WorldWide'}]}

        try:
            api = Finding(appid=self.api_key, 
                          config_file=None, 
                          siteid='EBAY-US')
            response = api.execute('findItemsAdvanced', payload)

            if response.reply.searchResult._count == '0':
                return

            key = excel.filterkey(key, 'PN')
            for item in response.reply.searchResult.item:
                title = item.title.upper().split()
                for title_key in title:
                    
                    title_key = excel.filterkey(title_key, 'PN')
                    if key in title_key:

                        price_item = round(float(item.sellingStatus.currentPrice.value))
                        api_get = Trading(appid=self.api_key, config_file=None, 
                                          certid=self.cert_id, devid= self.dev_id,
                                          token=self.token)

                        response_get_item = api_get.execute('GetItem', {'ItemID': item.itemId,
                                                                'IncludeItemSpecifics': True})

                        try:
                            for specifics in response_get_item.reply.Item.ItemSpecifics.NameValueList:
                                name = ('Model', 'MPN')
                                if specifics.Name in name:
                                
                                    filter_specific_key = excel.filterkey(specifics.Value, 'PN')
                                    if key in filter_specific_key or filter_specific_key in key:

                                        data[main_key]['URL'] = item.viewItemURL
                                        data[main_key]['СТОИМОСТЬ ТОВАРА/USD'] = price_item
                                        return
                        except:
                            continue
            return
        except:
            return


class Parse:

    """
    class for working with huawei buckets and searching for model/pn from the website
    -----
    * method find() allows you to search for model/pn on the huawei website using parsing
    """

    STATIC_URL: dict = {
        'HUAWEI': URL_HUAWEI
    }

    HEADERS: dict = {
        "User-Agent": USER_AGENT
    }

    def __init__(self, key: str) -> None:

        self.key = key
        self.url = self.STATIC_URL['HUAWEI'].format(key=key)
        self.session = HTMLSession()

    def find(self):
        """Find_model"""

        try:
            resp = self.session.get(self.url)
            resp.html.render()
            soup = BeautifulSoup(resp.html.html, "lxml")
            table = soup.find('table', id="model-table").find('tbody').find_all('tr')
            count = 0
            item_list = []

            for item in table:

                keys = ('Model', 'Part Number')
                try:
                    title = item.find(class_='con-title').text
                    if title in keys:
                        count += 1
                        value = item.find('pre').text
                        if self.key != value:
                            item_list.append(value)
                        if count == 2:
                            return item_list         
                except:
                    continue

        except:
            return        

