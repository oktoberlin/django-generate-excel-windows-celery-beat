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
    start_date = forms.DateTimeField(
        input_formats=['%Y-%m-%d %H:%M:%S'],
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input',
            'data-target': '#datetimepicker1'
        })
    )
    end_date = forms.DateTimeField(
        input_formats=['%Y-%m-%d %H:%M:%S'],
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input',
            'data-target': '#datetimepicker2'
        })
    )
    class Meta:
        model= generate_excel
        fields=['client','email','start_date','end_date','period',]