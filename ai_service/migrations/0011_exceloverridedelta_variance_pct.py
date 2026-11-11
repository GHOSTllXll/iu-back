# Feature 1 (Variance & Change Ledger) extension: adds
# ExcelOverrideDelta.variance_pct — percentage change between a delta's
# original_value and new_value, computed only when both sides were a real
# number (see _variance_pct() in views.py). Purely additive and nullable —
# every existing ExcelOverrideDelta row simply has no variance_pct, same
# "legacy rows don't get backfilled" precedent as 0010's own new fields.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0010_excel_override_audit'),
    ]

    operations = [
        migrations.AddField(
            model_name='exceloverridedelta',
            name='variance_pct',
            field=models.FloatField(blank=True, null=True),
        ),
    ]
