# Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): the one
# new table this feature needs on the ai_service side (the workspace
# address itself lives on users.TeamWorkspace — see
# users/migrations/0015_teamworkspace_inbound_email_slug.py). Purely
# additive; no backfill — there is no historical inbound email data to
# backfill from.
import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0017_lender_share_link'),
        ('users', '0015_teamworkspace_inbound_email_slug'),
    ]

    operations = [
        migrations.CreateModel(
            name='InboundEmailEvent',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('provider_message_id', models.CharField(max_length=255, unique=True)),
                ('sender_email', models.EmailField(blank=True, default='', max_length=254)),
                ('subject', models.CharField(blank=True, default='', max_length=500)),
                ('status', models.CharField(choices=[
                    ('accepted', 'Accepted'),
                    ('rejected_unknown_slug', 'Rejected — Unknown Workspace Slug'),
                    ('rejected_unauthorized_sender', 'Rejected — Unauthorized Sender'),
                    ('rejected_no_usable_attachments', 'Rejected — No Usable Attachments'),
                    ('rejected_malformed', 'Rejected — Malformed Payload'),
                ], max_length=32)),
                ('received_at', models.DateTimeField(auto_now_add=True)),
                ('project', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='inbound_email_events', to='ai_service.propertyproject')),
                ('workspace', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='inbound_email_events', to='users.teamworkspace')),
            ],
            options={
                'ordering': ['-received_at'],
            },
        ),
        migrations.AddIndex(
            model_name='inboundemailevent',
            index=models.Index(fields=['workspace', '-received_at'], name='inbound_email_workspace_idx'),
        ),
    ]
