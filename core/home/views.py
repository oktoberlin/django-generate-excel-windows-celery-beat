from django.shortcuts import render, redirect
from .forms import generate_excel_form

from .tasks import send_mail_task
# Create your views here.

def index(request):
    form = generate_excel_form(request.POST)
    if request.method == 'POST':
        # create a form instance and populate it with data from the request:
        #form = forms.daftar_pre_toefl_form(request.POST)
        # check whether it's valid:
        if form.is_valid():
            form.save()
            return redirect ("home:success")
    else:
        form = generate_excel_form()

    return render(request, 'main.html', {'form': form})


def success(request):
    return render(request, 'success.html')