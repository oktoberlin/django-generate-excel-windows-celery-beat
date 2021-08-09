from django.shortcuts import render, redirect
from .forms import generate_excel_form
from django.contrib.auth.decorators import login_required

from .tasks import send_mail_task
# Create your views here.

@login_required
def index(request):
    form = generate_excel_form(request.POST)
    if request.method == 'POST':
        # create a form instance and populate it with data from the request:
        #form = forms.daftar_pre_toefl_form(request.POST)
        # check whether it's valid:
        if form.is_valid():
            form.save()
            client = form.cleaned_data.get('client')
            email = form.cleaned_data.get('email')
            start_date = form.cleaned_data.get('start_date')
            end_date = form.cleaned_data.get('end_date')

            send_mail_task(client, email, start_date, end_date)
            return redirect ("home:success")
    else:
        form = generate_excel_form()

    return render(request, 'main.html', {'form': form})


def success(request):
    return render(request, 'success.html')