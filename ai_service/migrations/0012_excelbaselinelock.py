# Feature 2 (Immutable Snapshots): adds the ExcelBaselineLock table (the
# "Approve & Lock Baseline" record) plus
# ExcelOverrideUpload.is_post_lock_revision. Purely additive — the new
# table is brand new (nothing to backfill) and the new boolean field
# defaults to False for every existing ExcelOverrideUpload row, which is
# correct: none of them could have happened post-lock, since locking
# didn't exist yet.
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0011_exceloverridedelta_variance_pct'),
        ('users', '0013_customuser_workspace_role'),
    ]

    operations = [
        migrations.AddField(
            model_name='exceloverrideupload',
            name='is_post_lock_revision',
            field=models.BooleanField(default=False),
        ),
        migrations.CreateModel(
            name='ExcelBaselineLock',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('locked_at', models.DateTimeField(auto_now_add=True)),
                ('snapshot_sha256', models.CharField(max_length=64)),
                ('analysis', models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, related_name='baseline_locks', to='ai_service.analysisreport')),
                ('locked_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='locked_baselines', to=settings.AUTH_USER_MODEL)),
                ('project', models.OneToOneField(on_delete=django.db.models.deletion.CASCADE, related_name='baseline_lock', to='ai_service.propertyproject')),
            ],
        ),
    ]
