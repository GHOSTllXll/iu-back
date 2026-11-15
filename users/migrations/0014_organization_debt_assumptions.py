# Sprint "Workflow Upgrades" Feature 5 (Automatic Loan Sizing & DSCR
# Check): adds Organization.debt_assumptions — a single JSONField holding
# the org-wide defaults (target_ltv, interest_rate, amortization_years,
# min_dscr), editable by Org Admins only in Workspace Settings. Purely
# additive — defaults to {} for every existing organization, which
# Organization.get_debt_assumptions() then merges over
# DEFAULT_DEBT_ASSUMPTIONS at read time, so no backfill is needed.
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('users', '0013_customuser_workspace_role'),
    ]

    operations = [
        migrations.AddField(
            model_name='organization',
            name='debt_assumptions',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
