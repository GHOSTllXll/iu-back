# Sprint "Workflow Upgrades" Feature 1 (10-Second "Kill or Proceed"
# Scorecard & Missing Doc Check): adds PropertyProject.asking_price (the
# analyst's own entered/confirmed price, separate from the AI-extracted
# OM-stated price already inside AnalysisReport.metrics) and
# AnalysisReport.documents_provided (which of om/t12/rent_roll were
# actually submitted for a given run). Purely additive — both default to
# "nothing recorded yet" (null / {}) for every existing row, which reader
# code treats as the correct legacy behavior (see each field's own
# docstring in models.py) rather than needing a backfill migration.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0013_exceloverrideupload_variance_note'),
    ]

    operations = [
        migrations.AddField(
            model_name='propertyproject',
            name='asking_price',
            field=models.FloatField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='analysisreport',
            name='documents_provided',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
