from django.db import models

class Repair(models.Model):
    id = models.BigIntegerField(primary_key=True)
    contno = models.TextField(db_column='contNo', blank=True, null=True)  # Field name made lowercase.
    dmgloc = models.TextField(db_column='dmgLoc', blank=True, null=True)  # Field name made lowercase.
    compcode = models.TextField(db_column='compCode', blank=True, null=True)  # Field name made lowercase.
    dmgcode = models.TextField(db_column='dmgCode', blank=True, null=True)  # Field name made lowercase.
    actrpr = models.TextField(db_column='actRpr', blank=True, null=True)  # Field name made lowercase.
    length = models.FloatField(blank=True, null=True)
    width = models.FloatField(blank=True, null=True)
    percentage = models.FloatField(blank=True, null=True)
    rprqty = models.FloatField(db_column='rprQty', blank=True, null=True)  # Field name made lowercase.
    manhour = models.FloatField(blank=True, null=True)
    manhourdps = models.TextField(db_column='manhourDPS', blank=True, null=True)  # Field name made lowercase.
    matidr = models.FloatField(db_column='matIdr', blank=True, null=True)  # Field name made lowercase.
    materialdps = models.TextField(db_column='materialDPS', blank=True, null=True)  # Field name made lowercase.
    respparty = models.TextField(db_column='respParty', blank=True, null=True)  # Field name made lowercase.
    eorno = models.TextField(db_column='EORNo', blank=True, null=True)  # Field name made lowercase.
    rprcreate = models.DateTimeField(db_column='rprCreate', blank=True, null=True)  # Field name made lowercase.
    approve = models.BooleanField(default=False)

    class Meta:
        #managed = False
        db_table = 'repair'
'''
approve = models.BooleanField(default=False)
'''