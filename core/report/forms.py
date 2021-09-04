from report.models import report_Pdf
from django import forms
#from .models import generate_excel

class report_Pdf_form(forms.ModelForm):
    containerNumber=forms.CharField()
    
    class Meta:
        model= report_Pdf
        fields=['containerNumber',]