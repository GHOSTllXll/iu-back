# Feature 4 (Cell Override Audit Log): adds AnalysisReport.injected_values_snapshot
# plus the ExcelOverrideUpload/ExcelOverrideDelta tables. Purely additive —
# the new field is nullable (legacy/never-downloaded analyses simply have no
# snapshot, same precedent as rent_roll_data/t12_monthly_data) and both new
# tables are brand new, so there is no existing data to backfill and no
# RunPython step needed.
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0009_alter_analysisreport_project'),
        ('users', '0013_customuser_workspace_role'),
    ]

    operations = [
        migrations.AddField(
            model_name='analysisreport',
            name='injected_values_snapshot',
            field=models.JSONField(blank=True, null=True),
        ),
        migrations.CreateModel(
            name='ExcelOverrideUpload',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('original_filename', models.CharField(blank=True, default='', max_length=255)),
                ('uploaded_at', models.DateTimeField(auto_now_add=True)),
                ('analysis', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='override_uploads', to='ai_service.analysisreport')),
                ('project', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='override_uploads', to='ai_service.propertyproject')),
                ('uploaded_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='excel_override_uploads', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['-uploaded_at'],
            },
        ),
        migrations.AddIndex(
            model_name='exceloverrideupload',
            index=models.Index(fields=['project', '-uploaded_at'], name='override_upload_project_idx'),
        ),
        migrations.CreateModel(
            name='ExcelOverrideDelta',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('field_name', models.CharField(max_length=100)),
                ('row_index', models.IntegerField(blank=True, null=True)),
                ('original_value', models.TextField(blank=True, default='')),
                ('new_value', models.TextField(blank=True, default='')),
                ('upload', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='deltas', to='ai_service.exceloverrideupload')),
            ],
            options={
                'ordering': ['row_index', 'field_name'],
            },
        ),
    ]
