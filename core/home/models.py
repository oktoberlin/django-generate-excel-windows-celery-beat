from django.db import models

class generate_excel(models.Model):
    client=models.CharField(verbose_name="Client",max_length=30)
    email=models.EmailField(max_length=200)
