from django.db import migrations


def backfill_property_projects(apps, schema_editor):
    """
    Feature 3: auto-wraps every pre-existing AnalysisReport into its own
    single-analysis PropertyProject, per explicit product decision — no
    attempt is made to guess which historical analyses were actually
    re-runs of the same deal; each legacy row becomes its own project.

    Safe to re-run: only touches rows where project_id is still NULL, so a
    partial failure (and retry) never double-wraps a row that already got
    one.
    """
    AnalysisReport = apps.get_model('ai_service', 'AnalysisReport')
    PropertyProject = apps.get_model('ai_service', 'PropertyProject')

    unwrapped = AnalysisReport.objects.filter(project__isnull=True).iterator()
    for report in unwrapped:
        project = PropertyProject.objects.create(
            organization_id=report.organization_id,
            workspace=None,
            property_name=report.property_name,
            deal_status='reviewing',
            created_by_id=report.uploaded_by_id,
        )
        # created_at/updated_at are auto_now_add/auto_now — overridden here
        # via a queryset .update() rather than re-saving the instance, so the
        # project's history lines up with the analysis it was wrapped around
        # instead of "now" (the moment this migration happened to run).
        PropertyProject.objects.filter(pk=project.pk).update(
            created_at=report.created_at,
            updated_at=report.created_at,
        )

        AnalysisReport.objects.filter(pk=report.pk).update(project_id=project.pk)


def unwrap_property_projects(apps, schema_editor):
    """
    Reverse: only undoes wrapping performed by the forward function above —
    a PropertyProject with exactly one AnalysisReport and no Team Notes,
    which is the shape every project this migration creates has. A project
    that has since gained a second analysis (re-upload/re-analyze) or a
    Team Note is left alone, because by then it is real workflow data, not
    an artifact of this backfill, and un-wrapping it would lose that.
    """
    AnalysisReport = apps.get_model('ai_service', 'AnalysisReport')
    PropertyProject = apps.get_model('ai_service', 'PropertyProject')

    candidates = PropertyProject.objects.filter(analyses__isnull=False).distinct()

    for project in candidates.iterator():
        analysis_count = project.analyses.count()
        note_count = project.notes.count()
        if analysis_count == 1 and note_count == 0:
            AnalysisReport.objects.filter(project=project).update(project=None)
            project.delete()


class Migration(migrations.Migration):

    dependencies = [
        ('ai_service', '0007_analysisreport_rent_roll_data_and_more'),
    ]

    operations = [
        migrations.RunPython(backfill_property_projects, unwrap_property_projects),
    ]
