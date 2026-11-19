# Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): adds the
# globally-unique inbound_email_slug field to TeamWorkspace, then
# backfills every EXISTING workspace with one — consistent with this
# codebase's established pattern for a new field that every row needs a
# real value for (see 0008_backfill_propertyproject,
# 0013_customuser_workspace_role) rather than leaving old rows slug-less
# until they happen to be re-saved. Uses the exact same generation logic
# as TeamWorkspace._generate_unique_inbound_slug() in models.py, inlined
# here (not imported) per Django's own migration-writing guidance: a data
# migration must never depend on application code that can change
# later — only on the historical model state this migration operates on.
import uuid

from django.db import migrations, models
from django.utils.text import slugify


def _backfill_inbound_email_slugs(apps, schema_editor):
    TeamWorkspace = apps.get_model('users', 'TeamWorkspace')
    existing_slugs = set(
        TeamWorkspace.objects.exclude(inbound_email_slug__isnull=True).values_list('inbound_email_slug', flat=True)
    )

    for workspace in TeamWorkspace.objects.filter(inbound_email_slug__isnull=True):
        base = slugify(f"{workspace.organization_id}-{workspace.name}")[:40] or 'workspace'
        candidate = None
        for _ in range(5):
            attempt = f"{base}-{uuid.uuid4().hex[:6]}"
            if attempt not in existing_slugs:
                candidate = attempt
                break
        if candidate is None:
            candidate = uuid.uuid4().hex
        existing_slugs.add(candidate)
        workspace.inbound_email_slug = candidate
        workspace.save(update_fields=['inbound_email_slug'])


def _noop_reverse(apps, schema_editor):
    # Reversing this migration only needs to undo the AddField below
    # (handled automatically); there's nothing meaningful to "unbackfill" —
    # see 0008_backfill_propertyproject for the same reverse-is-a-noop
    # reasoning applied to an earlier backfill in this app.
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('users', '0014_organization_debt_assumptions'),
    ]

    operations = [
        migrations.AddField(
            model_name='teamworkspace',
            name='inbound_email_slug',
            field=models.SlugField(blank=True, editable=False, max_length=80, null=True, unique=True),
        ),
        migrations.RunPython(_backfill_inbound_email_slugs, _noop_reverse),
    ]
