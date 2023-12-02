from django.http import HttpResponse
from django.http import HttpResponseRedirect

def code_view(request):
    url = request.GET.get('url')

    if url is None:
        return HttpResponse('No URL specified')
    
    return HttpResponseRedirect(url)
