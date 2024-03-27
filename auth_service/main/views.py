import requests
from django.http import HttpResponse
from django.views import generic


class SendEgovAuthRequest(generic.View):
    def get(self, request):
        return HttpResponse('Hello, World!')

    def post(self, request):
        print("request POST:", request.POST)
        data = {
            'certificate': request.POST.get('certificate')
        }
        print("data:", data)
        try:
            response = requests.post(
                url='https://idp.egov.kz/idp/eds-login.do/',
                data=data,
                headers={
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            )
            if response.status_code != 200:
                raise Exception(response)
            return response
        except Exception as e:
            raise Exception(f'Error occurred while sending request: {e}')
