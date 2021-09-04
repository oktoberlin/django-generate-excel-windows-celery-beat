from django.db import models

class Test(models.Model):
    index = models.BigIntegerField(blank=True, null=True)
    contno = models.TextField(db_column='contNo', blank=True, null=True)  # Field name made lowercase.
    dmgloc = models.TextField(db_column='dmgLoc', blank=True, null=True)  # Field name made lowercase.
    dmgcode = models.TextField(db_column='dmgCode', blank=True, null=True)  # Field name made lowercase.
    compcode = models.TextField(db_column='compCode', blank=True, null=True)  # Field name made lowercase.
    width = models.FloatField(blank=True, null=True)
    lebar = models.FloatField(blank=True, null=True)
    percentage = models.FloatField(blank=True, null=True)
    rprqty = models.FloatField(db_column='rprQty', blank=True, null=True)  # Field name made lowercase.
    manhour = models.FloatField(blank=True, null=True)
    matidr = models.FloatField(db_column='matIdr', blank=True, null=True)  # Field name made lowercase.
    actrpr = models.TextField(db_column='actRpr', blank=True, null=True)  # Field name made lowercase.
    respparty = models.TextField(db_column='respParty', blank=True, null=True)  # Field name made lowercase.
    eorno = models.TextField(db_column='EORNo', blank=True, null=True)  # Field name made lowercase.
    rprcreate = models.DateTimeField(db_column='rprCreate', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'test'