# Feature 4 ("Head of Underwriting" hierarchy): adds CustomUser.workspace_role.
#
# Purely additive — no existing column is altered or tightened. The one
# thing worth explaining is the backfill:
#
# The field's own default is 'analyst' (the safer choice for every NEW
# workspace assignment from here on). But BEFORE this field existed, every
# member of a workspace already had full visibility into that workspace's
# whole deal pipeline (see the old visible_projects_for behavior in
# ai_service/permissions.py, prior to this feature). If we let every
# existing workspace member default to 'analyst', this migration would
# silently SHRINK their visibility the moment it's deployed — a production
# behavior regression with no code change and no one asking for it.
#
# So: any user who ALREADY has a workspace assigned as of this migration
# keeps today's behavior exactly by being backfilled to 'head'. Only
# brand-new workspace assignments going forward start at the safer
# 'analyst' default and need an explicit promotion.
from django.db import migrations, models


def backfill_existing_workspace_members_to_head(apps, schema_editor):
    CustomUser = apps.get_model('users', 'CustomUser')
    CustomUser.objects.filter(workspace_id__isnull=False).update(workspace_role='head')


def noop_reverse(apps, schema_editor):
    # Nothing to undo — reversing this migration just drops the column
    # (handled by the RemoveField Django generates automatically), and
    # un-backfilling workspace_role has no meaning once the column is gone.
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('users', '0012_teamworkspace_customuser_workspace_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='customuser',
            name='workspace_role',
            field=models.CharField(
                choices=[('analyst', 'Analyst'), ('head', 'Head of Underwriting')],
                default='analyst',
                max_length=10,
            ),
        ),
        migrations.RunPython(backfill_existing_workspace_members_to_head, noop_reverse),
    ]
