import base64
import requests


def sign_token(in_data, base64_sign, password, address):
    nca_node_address = address
    xml_sign_request = {
        'xml': in_data,
        'signers': [
            {
                'key': base64_sign,
                'password': password,
                'keyAlias': None
            }
        ],
        'clearSignatures': False,
        'trimXml': False
    }
    try:
        response = requests.post(f'{nca_node_address}/xml/sign', json=xml_sign_request)
        signed_xml = response.json()
        if signed_xml['status'] != 200:
            raise Exception(signed_xml)
        return signed_xml
    except Exception as e:
        raise Exception(f'Error occurred while signing XML: {e}')


def send_egov_auth_request(result):
    data = {
        'certificate': result['xml']
    }

    try:
        response = requests.get(
            url='https://idp.egov.kz/idp/eds-login.do/',
            data=data
        )
        if response.status_code != 200:
            raise Exception(response)
        return response

    except Exception as e:
        raise Exception(f'Error occurred while sending request: {e}')


def send_auth_request(sessionID, timeTicket):
    in_data = f'<?xml version=\"1.0\" encoding=\"utf-8\"?><login><timeTicket>{timeTicket}</timeTicket><sessionid>{sessionID}</sessionid></login>'
    with open('eds/AUTH_RSA256_f76176815e5e30559813d90ab8aff7b0219b1795.p12', 'rb') as file:
        content = file.read()
        base64_sign = base64.b64encode(content).decode('utf-8')
    password = "May2021"
    address = "http://localhost:14579"

    result = sign_token(in_data, base64_sign, password, address)

    return send_egov_auth_request(result)
