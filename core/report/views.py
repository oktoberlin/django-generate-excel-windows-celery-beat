from django.shortcuts import render, redirect
#from .forms import Tes
from django.contrib.auth.decorators import login_required
from .models import Test
from django.views.decorators.csrf import csrf_exempt
#from .tasks import send_mail_task
# Create your views here.

#@login_required
#@csrf_exempt
'''
def index(request):
    ContainerNumber=None
    if request.GET.get('contNo'):
        contNo = request.GET.get('contNo')
        ContainerNumber = Test.objects.filter(query__icontains=contNo)
        name = request.GET.get('name')
        query = Test.object.create(query=contNo, user_id=name)
        query.save()
    return render(request, 'main_report.html',{
        'containerNumber': ContainerNumber,
    })
'''
def index(request):
    data = Test.objects.get('contNo')
    context = {
        "containers": data
    }
    return render(request, 'main_report.html',context)

def success(request):
    return render(request, 'success_report.html')