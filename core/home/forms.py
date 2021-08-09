from django import forms
from .models import generate_excel
OPTIONS = (
        ("D","Daily"),
        ("M","Monthly"),
        ("A","Annually"),
    )
class generate_excel_form(forms.ModelForm):
    client=forms.CharField()
    email=forms.EmailField()
    period = forms.MultipleChoiceField(widget=forms.CheckboxSelectMultiple,
                                          choices=OPTIONS)
    start_date = forms.CharField()
    end_date = forms.CharField()
    class Meta:
        model= generate_excel
        fields=['client','email','start_date','end_date','period',]