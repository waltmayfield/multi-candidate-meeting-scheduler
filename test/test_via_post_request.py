
#https://www.w3schools.com/python/ref_requests_post.asp

import requests

url = 'https://ballard.amazon.com/owa/service.svc'
headers = {
	    action: 'FindPeople',
	    'content-type': 'application/json; charset=UTF-8',
	    'x-owa-actionname': 'OwaOptionPage',
	    'x-owa-canary': this.getToken()#Fix this part
	}

# data = { 'Header':
#  data = """
# 	    Header: {
# 	        RequestServerVersion: 'Exchange2013'
# 	    },
# 	    Body: {
# 	        IndexedPageItemView: {
# 	            __type: 'IndexedPageView:#Exchange',
# 	            BasePoint: 'Beginning'
# 	        },
# 	        QueryString: 'waltmayf@'
# 	    }
# }
# """

data = '{"Header":{"RequestServerVersion":"Exchange2013"},"Body":{"IndexedPageItemView":{"__type":"IndexedPageView:#Exchange","BasePoint":"Beginning"},"QueryString":"waltmayf@"}}'

x = requests.post(url, data = myobj)

print(x.text)