# Pipeline Modules Phase 8 (Live In-Browser Spreadsheet Editor): adds
# 'live_grid_edit' to DealVaultFile.file_type's choices so the Live Grid
# Editor's "Save" action (GridEditSaveView in views.py) can vault its
# synthetic diff workbook under its own honest label rather than being
# forced into the pre-existing 're_upload' choice. Choices are
# application-level validation only (no database constraint changes for
# either SQLite or MySQL's VARCHAR-backed CharField here), so this
# migration only updates Django's own migration state — nothing to
# backfill, nothing to reverse beyond the field state itself.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0018_inbound_email_event'),
    ]

    operations = [
        migrations.AlterField(
            model_name='dealvaultfile',
            name='file_type',
            field=models.CharField(choices=[
                ('locked_baseline_excel', 'Locked Baseline Excel'),
                ('re_upload', 'Re-Uploaded Excel'),
                ('supporting_document', 'Supporting Document'),
                ('live_grid_edit', 'Live Grid Edit'),
            ], max_length=24),
        ),
    ]
