# V2 Module 3 (Live Debt & Equity Waterfall Scenario Builder): two purely
# additive JSON fields — AnalysisReport.waterfall_scenario (the working,
# editable scenario) and ExcelBaselineLock.waterfall_summary (a frozen
# copy captured at lock time). See each field's own docstring in
# models.py. No backfill — there is no historical scenario data.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0019_dealvaultfile_live_grid_edit_choice'),
    ]

    operations = [
        migrations.AddField(
            model_name='analysisreport',
            name='waterfall_scenario',
            field=models.JSONField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='excelbaselinelock',
            name='waterfall_summary',
            field=models.JSONField(blank=True, null=True),
        ),
    ]
