# backend/ai_service/permissions.py
"""
Feature 3 (Enterprise Pipeline & Workflow Dashboard) / Feature 4 ("Head of
Underwriting" hierarchy): visibility rules for PropertyProject, shared by
every Deal Pipeline endpoint (list, detail, notes, download, override-check)
so "who can see this deal" is defined in exactly one place rather than
re-derived per view.

Visibility rule, exactly as approved:
  - Platform superuser (CustomUser.is_admin)             -> every project, any org.
  - Org admin (CustomUser.role == 'admin')                -> every project in
    their own organization, regardless of workspace.
  - Workspace Head (CustomUser.workspace_role == 'head',
    requires a workspace assigned)                        -> every project in
    THEIR OWN workspace only — confirmed never cross-workspace; only an org
    admin gets that.
  - Everyone else — a plain workspace member, OR a member
    with NO workspace assigned (the expected state during
    rollout, before an org admin has set workspaces up,
    or for any account created before this feature)        -> only projects
    THEY personally created - a safe fallback that's neither "see
    nothing" nor "see the whole org" (and, before Feature 4 existed, this
    USED to also be true for every plain workspace member automatically
    getting their whole workspace's visibility — see
    users/migrations/0013_customuser_workspace_role.py for why every
    user who already had a workspace before this field existed was
    backfilled to workspace_role='head', so THIS deploy doesn't quietly
    shrink anyone's existing visibility; only a NEW workspace assignment
    from here on starts at the tighter 'analyst' default).
"""
from .models import PropertyProject


def visible_projects_for(user):
    """
    Returns a PropertyProject QuerySet - never a list - so callers can keep
    chaining .filter()/.order_by()/pagination on top rather than
    materializing early.
    """
    base = PropertyProject.objects.filter(is_deleted=False)

    if not getattr(user, 'is_authenticated', False):
        return base.none()

    if user.is_admin:
        return base

    if user.organization_id is None:
        # Mirrors the pattern used elsewhere (e.g.
        # MyOrganizationExcelTemplateView) - an account with no organization
        # has nothing of its own to see.
        return base.none()

    org_scoped = base.filter(organization_id=user.organization_id)

    if user.role == 'admin':
        return org_scoped

    if user.workspace_id is not None and user.workspace_role == 'head':
        return org_scoped.filter(workspace_id=user.workspace_id)

    return org_scoped.filter(created_by_id=user.pk)


def can_manage_project(user, project) -> bool:
    """
    Whether `user` may change this ONE already-fetched project (update its
    deal_status, add a Team Note, download its Excel). Re-derives the same
    rule as visible_projects_for rather than re-querying it, since the
    caller already has the instance in hand from a prior lookup.
    """
    if not getattr(user, 'is_authenticated', False):
        return False
    if user.is_admin:
        return True
    if user.organization_id is None or project.organization_id != user.organization_id:
        return False
    if user.role == 'admin':
        return True
    if user.workspace_id is not None and user.workspace_role == 'head':
        return project.workspace_id == user.workspace_id
    return project.created_by_id == user.pk
