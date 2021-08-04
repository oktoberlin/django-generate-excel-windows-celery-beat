from django import forms
from .models import generate_excel

class generate_excel_form(forms.ModelForm):
    client=forms.CharField()
    email=forms.EmailField()
   
    class Meta:
        model= generate_excel
        fields=['client','email',]