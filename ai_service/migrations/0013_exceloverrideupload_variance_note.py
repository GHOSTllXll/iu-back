# Sprint "Workflow Upgrades" Feature 2 (Variance Alert Context Notes): adds
# ExcelOverrideUpload.variance_note — a required-at-the-view-layer (not at
# the DB layer) 1-sentence leadership-facing reason, collected only when an
# upload's deltas cross the existing >2.0% GPR/OpEx/NOI threshold (see
# OVERRIDE_ALERT_TRACKED_FIELDS / OVERRIDE_ALERT_THRESHOLD_PCT and
# ExcelOverrideCheckView.post() in views.py). Purely additive and blank by
# default — every existing upload simply has no note, same "legacy rows
# aren't backfilled" precedent as 0011/0012's new fields.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0012_excelbaselinelock'),
    ]

    operations = [
        migrations.AddField(
            model_name='exceloverrideupload',
            name='variance_note',
            field=models.TextField(blank=True, default=''),
        ),
    ]
