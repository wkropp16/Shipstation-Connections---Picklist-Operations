import requests, json

def getShipStationHttpRequest(reference,params): #HTTP Request for ShipStation API. Returns a JSON formatted file.
	url = 			('https://ssapi.shipstation.com/' + str(reference) + '/')
	params = 		('?' + str(params))
	headers = 		{'Authorization': 'Basic REMOVED'}
	combinedURL = 	(''.join([url, params]))
	r = requests.get(combinedURL, headers=headers)
	r.raise_for_status()
	return r;

def postShipStationHttpRequest(reference,params,payload): #HTTP Request for ShipStation API. Returns a JSON formatted file.
	url = 			('https://ssapi.shipstation.com/' + str(reference) + '/')
	params = 		(str(params))
	headers = 		{'Content-Type': 'application/json','Authorization': 'Basic REMOVED'}
	combinedURL = 	(''.join([url, params]))
	r = requests.post(combinedURL, data=json.dumps(payload), headers=headers)
	r.raise_for_status()
	return r;

