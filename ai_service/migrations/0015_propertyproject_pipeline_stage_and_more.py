# Pipeline Modules Phase 1 (Daily Deal Pipeline & Team Command Center):
#
# 1. Replaces PropertyProject's old 3-state STATUS_CHOICES/deal_status
#    (reviewing/approved/archived) with the new 5-stage
#    PIPELINE_STAGE_CHOICES/pipeline_stage — per explicit client decision,
#    a clean replacement rather than a second parallel field, so there's
#    zero status confusion on the deal cards going forward. The old field
#    was a business OUTCOME; the new one is a workflow funnel POSITION —
#    different axes, so this is a one-time, best-effort value mapping,
#    not a lossless rename:
#      reviewing -> underwriting
#      approved  -> offer_submitted
#      archived  -> dead_passed
#    Applied via RunPython between adding the new field and removing the
#    old one, so every existing row keeps a sensible stage rather than
#    silently resetting to the new field's own default.
# 2. Adds PropertyProject.assigned_analyst (who's actively working the
#    deal day to day) and target_closing_or_offer_date (the team's own
#    planning date) — both purely additive/nullable, nothing to backfill.
# 3. Adds the new DealActivityLog table (stage-transition history) — see
#    that model's own docstring in models.py for why it's a different
#    thing from users.ActivityLog and from ExcelOverrideUpload/-Delta.
#    Brand new table, nothing to backfill; existing deals simply have no
#    activity history before this migration runs (their very next stage
#    change is the first row they get).
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


def migrate_deal_status_to_pipeline_stage(apps, schema_editor):
    PropertyProject = apps.get_model('ai_service', 'PropertyProject')
    STAGE_MAP = {
        'reviewing': 'underwriting',
        'approved': 'offer_submitted',
        'archived': 'dead_passed',
    }
    for old_value, new_value in STAGE_MAP.items():
        PropertyProject.objects.filter(deal_status=old_value).update(pipeline_stage=new_value)


def reverse_pipeline_stage_to_deal_status(apps, schema_editor):
    # Best-effort only — several pipeline_stage values collapse onto the
    # SAME deal_status (new_incoming/underwriting/vp_review all have no
    # clean pre-image other than 'reviewing'), so this is lossy by
    # necessity. Only ever exercised if this migration is unapplied.
    PropertyProject = apps.get_model('ai_service', 'PropertyProject')
    REVERSE_MAP = {
        'new_incoming': 'reviewing',
        'underwriting': 'reviewing',
        'vp_review': 'reviewing',
        'offer_submitted': 'approved',
        'dead_passed': 'archived',
    }
    for old_value, new_value in REVERSE_MAP.items():
        PropertyProject.objects.filter(pipeline_stage=old_value).update(deal_status=new_value)


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0014_propertyproject_asking_price_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='propertyproject',
            name='pipeline_stage',
            field=models.CharField(
                choices=[
                    ('new_incoming', 'New Incoming'),
                    ('underwriting', 'Underwriting'),
                    ('vp_review', 'VP Review'),
                    ('offer_submitted', 'Offer Submitted'),
                    ('dead_passed', 'Dead / Passed'),
                ],
                default='new_incoming',
                max_length=20,
            ),
        ),
        migrations.RunPython(
            migrate_deal_status_to_pipeline_stage,
            reverse_pipeline_stage_to_deal_status,
        ),
        migrations.RemoveField(
            model_name='propertyproject',
            name='deal_status',
        ),
        migrations.AddField(
            model_name='propertyproject',
            name='assigned_analyst',
            field=models.ForeignKey(
                blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL,
                related_name='assigned_deals', to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddField(
            model_name='propertyproject',
            name='target_closing_or_offer_date',
            field=models.DateField(blank=True, null=True),
        ),
        migrations.CreateModel(
            name='DealActivityLog',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('from_stage', models.CharField(
                    blank=True, null=True, max_length=20,
                    choices=[
                        ('new_incoming', 'New Incoming'),
                        ('underwriting', 'Underwriting'),
                        ('vp_review', 'VP Review'),
                        ('offer_submitted', 'Offer Submitted'),
                        ('dead_passed', 'Dead / Passed'),
                    ],
                )),
                ('to_stage', models.CharField(
                    max_length=20,
                    choices=[
                        ('new_incoming', 'New Incoming'),
                        ('underwriting', 'Underwriting'),
                        ('vp_review', 'VP Review'),
                        ('offer_submitted', 'Offer Submitted'),
                        ('dead_passed', 'Dead / Passed'),
                    ],
                )),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('project', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='activity_log', to='ai_service.propertyproject')),
                ('user', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='deal_stage_transitions', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['-created_at'],
            },
        ),
        migrations.AddIndex(
            model_name='dealactivitylog',
            index=models.Index(fields=['project', '-created_at'], name='deal_activity_project_idx'),
        ),
    ]
