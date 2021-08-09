from django.contrib.auth.models import User
from django.contrib.auth import get_user_model
UserModel = get_user_model()
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.tokens import default_token_generator
from django.contrib.sites.shortcuts import get_current_site
from django.core.mail import EmailMessage
from django.http import HttpResponse

from django.shortcuts import render, redirect

from django.template.loader import render_to_string
from django.utils.encoding import force_bytes
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode


from django.contrib import messages
from django.contrib.auth.decorators import login_required
from .models import *
from .forms import UserUpdateForm, ProfileUpdateForm



def loginPage(request):
	if request.user.is_authenticated:
		return redirect('home:home')
	else:
		if request.method == 'POST':
			email = request.POST.get('email')
			password =request.POST.get('password')

			user = authenticate(request, email=email, password=password)

			if user is not None:
				login(request, user)
				return redirect('home:home')
			else:
				messages.info(request, 'Email atau Password Salah')

		context = {}
		return render(request, 'login.html', context)


def logoutUser(request):
	logout(request)
	return redirect('community:community')
@login_required
def profile(request):
    if request.method == 'POST':
        u_form = UserUpdateForm(request.POST, instance=request.user)
        p_form = ProfileUpdateForm(
            request.POST, request.FILES, instance=request.user.profile)

        if u_form.is_valid() and p_form.is_valid():
            u_form.save()
            p_form.save()
            messages.success(
                request, f'Akun Anda sukses diupdate')
            return redirect('profile')
    else:
        u_form = UserUpdateForm(instance=request.user)
        p_form = ProfileUpdateForm(instance=request.user.profile)
    context = {
        'u_form': u_form,
        'p_form': p_form
    }
    return render(request, 'profile.html', context)
