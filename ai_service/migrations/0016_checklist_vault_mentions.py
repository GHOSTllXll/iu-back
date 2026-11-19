# Pipeline Modules Phase 4 (Activity Feed, Comments, @Mentions, Persistent
# Missing-Doc Checklist) & Phase 5 (Excel Versioning & Supporting Document
# Vault) — ships both phases' schema together in one migration, per the
# client's explicit "finish this entire pipeline module cleanly" request.
# All three changes are purely additive — nothing to backfill:
#
# 1. TeamNote.mentioned_users (M2M to CustomUser): @mentions are in-app
#    only (no email infrastructure exists in this codebase yet) — see
#    that field's own docstring in models.py. Every existing TeamNote
#    simply has zero mentions, which is correct (none were ever parsed
#    before this field existed).
# 2. New DealChecklistItem table — the persistent missing-doc checklist,
#    seeded going forward from documents_provided (see
#    record_analysis_report() in document_history.py) plus whatever the
#    team adds by hand. Brand new table; no historical backfill attempted
#    — a deal's checklist simply starts empty as of this migration and
#    gets seeded the next time it's (re-)analyzed.
# 3. New DealVaultFile table — the Excel-versioning/supporting-document
#    vault (see that model's own docstring for exactly when a version is
#    auto-captured). Brand new table; existing deals simply have no
#    vaulted files until their next lock/re-upload/manual add.
import ai_service.models
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0015_propertyproject_pipeline_stage_and_more'),
        ('users', '0014_organization_debt_assumptions'),
    ]

    operations = [
        migrations.AddField(
            model_name='teamnote',
            name='mentioned_users',
            field=models.ManyToManyField(blank=True, related_name='deal_mentions', to=settings.AUTH_USER_MODEL),
        ),
        migrations.CreateModel(
            name='DealChecklistItem',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('doc_key', models.CharField(
                    blank=True, null=True, max_length=20,
                    choices=[
                        ('om', 'Offering Memorandum (OM)'),
                        ('t12', 'T12 Financial Statement'),
                        ('rent_roll', 'Rent Roll'),
                    ],
                )),
                ('label', models.CharField(max_length=255)),
                ('is_resolved', models.BooleanField(default=False)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('resolved_at', models.DateTimeField(blank=True, null=True)),
                ('project', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='checklist_items', to='ai_service.propertyproject')),
                ('created_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='created_checklist_items', to=settings.AUTH_USER_MODEL)),
                ('resolved_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='resolved_checklist_items', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['is_resolved', 'created_at'],
            },
        ),
        migrations.AddIndex(
            model_name='dealchecklistitem',
            index=models.Index(fields=['project', 'is_resolved'], name='checklist_project_resolved_idx'),
        ),
        migrations.CreateModel(
            name='DealVaultFile',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('file', models.FileField(storage=ai_service.models.deal_vault_storage, upload_to=ai_service.models.deal_vault_upload_path)),
                ('file_type', models.CharField(
                    max_length=24,
                    choices=[
                        ('locked_baseline_excel', 'Locked Baseline Excel'),
                        ('re_upload', 'Re-Uploaded Excel'),
                        ('supporting_document', 'Supporting Document'),
                    ],
                )),
                ('original_filename', models.CharField(max_length=255)),
                ('version_label', models.CharField(blank=True, default='', max_length=255)),
                ('uploaded_at', models.DateTimeField(auto_now_add=True)),
                ('project', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='vault_files', to='ai_service.propertyproject')),
                ('uploaded_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='vaulted_deal_files', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['-uploaded_at'],
            },
        ),
        migrations.AddIndex(
            model_name='dealvaultfile',
            index=models.Index(fields=['project', '-uploaded_at'], name='deal_vault_project_idx'),
        ),
    ]
