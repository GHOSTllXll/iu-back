# Pipeline Modules Phase 6 (One-Click Lender Submission Packet & View
# Tracking) — ships the one new table this feature needs. Purely
# additive; no backfill — every deal simply has zero LenderShareLink rows
# until someone generates one for it.
import ai_service.models
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0016_checklist_vault_mentions'),
        ('users', '0014_organization_debt_assumptions'),
    ]

    operations = [
        migrations.CreateModel(
            name='LenderShareLink',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('access_token', models.CharField(
                    default=ai_service.models._generate_lender_share_token,
                    editable=False, max_length=64, unique=True,
                )),
                ('recipient_label', models.CharField(blank=True, default='', max_length=255)),
                ('is_active', models.BooleanField(default=True)),
                ('view_count', models.PositiveIntegerField(default=0)),
                ('last_viewed_at', models.DateTimeField(blank=True, null=True)),
                ('expires_at', models.DateTimeField(blank=True, null=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('project', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='lender_share_links', to='ai_service.propertyproject')),
                ('workspace', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='lender_share_links', to='users.teamworkspace')),
                ('created_by', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='created_lender_share_links', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['-created_at'],
            },
        ),
        migrations.AddIndex(
            model_name='lendersharelink',
            index=models.Index(fields=['project', '-created_at'], name='lender_share_project_idx'),
        ),
    ]
