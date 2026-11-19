# users/models.py
import uuid
from django.conf import settings as django_settings
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models
from django.utils import timezone
from django.utils.text import slugify
from datetime import timedelta

class OrganizationManager(models.Manager):
    pass

class Organization(models.Model):
    PLAN_CHOICES = [
        ('basic', 'Basic'),
        ('professional', 'Professional'),
        ('enterprise', 'Enterprise'),
        ('trial', 'Trial')
    ]
    STATUS_CHOICES = [
        ('active', 'Active'),
        ('pending', 'Pending'),
        ('inactive', 'Inactive'),
    ]

    is_trial = models.BooleanField(default=False)
    trial_expires_at = models.DateTimeField(null=True, blank=True)
    # is_trial + trial_expires_at together drive the temp-account lifecycle:
    # a trial org's user gets exactly 1 analysis (via TIER_UPLOAD_LIMITS —
    # see ai_service/views.py), and the whole org+user gets hard-deleted by
    # the cleanup_expired_trials management command once trial_expires_at
    # passes, regardless of whether they ever used their one analysis.

    name = models.CharField(max_length=255, unique=True)
    primary_contact_name = models.CharField(max_length=100)
    primary_contact_surname = models.CharField(max_length=100)
    primary_email = models.EmailField(unique=True)
    
    subscription_plan = models.CharField(max_length=20, choices=PLAN_CHOICES, default='basic')
    team_size_limit = models.IntegerField(default=1)
    is_paid = models.BooleanField(default=False)
    account_status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    date_created = models.DateTimeField(auto_now_add=True)

    # Informational / bookkeeping only — e.g. "colliers.com" for an
    # institutional client. Org creation is entirely admin-provisioned (see
    # AdminCreateUserView / AdminCreateTrialAccountView), not self-serve
    # signup, so this is NOT used to look up or auto-assign an organization
    # at login — request.user.organization already does that directly.
    # This exists purely so the admin org list is easy to scan/search by
    # client domain, and as a human-readable label next to a custom Excel
    # template (see ai_service.models.ClientExcelTemplate).
    email_domain = models.CharField(max_length=255, blank=True, default='')

    # Institutional PPM feature gate — separate from subscription_plan on
    # purpose, so PPM access can be granted/revoked independently of tier
    # without needing a new PLAN_CHOICES value.
    has_ppm_access = models.BooleanField(default=False)

    # Per-org override for the rolling 30-day upload ceiling (see
    # TIER_UPLOAD_LIMITS / check_upload_quota in ai_service/views.py). Null
    # means "no override — use the standard tier-based limit". Set to an
    # integer for bespoke arrangements (e.g. a custom negotiated quota)
    # without needing a new tier or a hardcoded per-org special case.
    quota_override = models.IntegerField(null=True, blank=True)

    # Sprint "Workflow Upgrades" Feature 5 (Automatic Loan Sizing & DSCR
    # Check): org-wide defaults for debt sizing, editable by Org Admins
    # only (see DebtAssumptionsView in ai_service/views.py). A single
    # JSONField rather than four separate typed columns — deliberately the
    # first generic per-org settings bucket on this model; every existing
    # column above is a specific typed field, so this is the new pattern
    # for anything configurable-but-not-yet-numerous enough to earn its
    # own migration every time a client wants one more knob.
    # Keys: target_ltv (pct, e.g. 70.0), interest_rate (pct, e.g. 6.50),
    # amortization_years (int, e.g. 30), min_dscr (e.g. 1.25).
    # default=dict (not a lambda/callable literal) is Django's required
    # pattern for a mutable default on a JSONField — each row gets its own
    # fresh dict, never a shared one.
    debt_assumptions = models.JSONField(default=dict, blank=True)

    DEFAULT_DEBT_ASSUMPTIONS = {
        'target_ltv': 70.0,
        'interest_rate': 6.50,
        'amortization_years': 30,
        'min_dscr': 1.25,
    }

    def get_debt_assumptions(self):
        """
        Merges self.debt_assumptions over DEFAULT_DEBT_ASSUMPTIONS rather
        than returning the raw field — an org that has never touched
        Workspace Settings has debt_assumptions == {} (the JSONField
        default), and a partially-filled-in dict (e.g. only min_dscr ever
        PATCHed) should still fall back to the defaults for every other
        key, not silently compute with a missing one.
        """
        merged = dict(self.DEFAULT_DEBT_ASSUMPTIONS)
        merged.update(self.debt_assumptions or {})
        return merged

    objects = OrganizationManager()

    def __str__(self):
        return self.name


class TeamWorkspace(models.Model):
    """
    Feature 3 (Enterprise Pipeline & Workflow Dashboard): the middle tier of
    Company Account -> Team Workspace -> Property Project. Purely additive —
    a brand-new table, no existing data or behavior touched by its creation.

    A workspace belongs to exactly one Organization. A CustomUser belongs to
    at most one TeamWorkspace (see CustomUser.workspace below) — not a
    many-to-many membership table, since nothing in this product currently
    needs a person to sit in more than one workspace at once; that's a real
    product decision to make deliberately later; not something to
    half-support now by guessing.

    Deliberately does NOT itself enforce that a user's workspace belongs to
    that same user's organization — Django can't express a constraint across
    two different FKs at the database level. That invariant is enforced at
    the application layer instead (serializer validation on CustomUser),
    the same way CustomUser.save() already enforces is_admin at the
    application layer rather than the database.
    """
    organization = models.ForeignKey(Organization, on_delete=models.CASCADE, related_name='workspaces')
    name = models.CharField(max_length=255)

    created_by = models.ForeignKey(
        'CustomUser', on_delete=models.SET_NULL, null=True, blank=True, related_name='created_workspaces'
    )
    created_at = models.DateTimeField(auto_now_add=True)

    # Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): the
    # local-part half of this workspace's dedicated inbound address (e.g.
    # 'acme-dispositions-a1b2c3' -> deals-acme-dispositions-a1b2c3@
    # inboundunderwriting.com — see inbound_email_address below). Must be
    # GLOBALLY unique, not just unique per-org like `name` above — unlike
    # every other uniqueness constraint on this model, the inbound email
    # webhook (InboundEmailWebhookView in ai_service/views.py) has to
    # resolve a workspace from the recipient address ALONE, with no
    # organization context yet to scope the lookup by. Auto-generated
    # once, in save() below, on first creation — never user-editable
    # (there's no rename endpoint for it; renaming the WORKSPACE's
    # display `name` deliberately leaves this alone, so an address once
    # handed out to a lender/broker never silently breaks).
    inbound_email_slug = models.SlugField(max_length=80, unique=True, null=True, blank=True, editable=False)

    # V2 Module 4a (Buy-Box Auto-Alerts): per-workspace match criteria for
    # the "new deal matches our buy-box" alert, editable by this
    # workspace's Org Admin or Workspace Head only (see
    # WorkspaceBuyBoxView in ai_service/views.py) — exactly the same
    # JSONField-with-merged-defaults pattern as Organization.
    # debt_assumptions/get_debt_assumptions() above, applied one level
    # down the hierarchy since buy-box criteria is naturally a per-
    # workspace (not firm-wide) concept — different workspaces
    # (dispositions teams, asset classes) chase different deals. See
    # DEFAULT_BUY_BOX_CRITERIA/get_buy_box_criteria() below for the keys.
    buy_box_criteria = models.JSONField(default=dict, blank=True)

    DEFAULT_BUY_BOX_CRITERIA = {
        'enabled': False,
        # Matched case-insensitively against property_metadata.asset_class
        # (e.g. "Class B", "Class C"); an empty list means "any class".
        'asset_classes': [],
        'min_units': None,
        'max_units': None,
        'min_cap_rate_pct': None,
        'max_cap_rate_pct': None,
        'max_price_per_unit': None,
        # Matched as a case-insensitive substring/suffix against the
        # deal's full address string (see buy_box._location_matches) —
        # this app doesn't store a structured city/state at the
        # property level, only the AI-extracted free-text address.
        'target_cities': [],
        'target_states': [],
    }

    def get_buy_box_criteria(self):
        """
        Merges self.buy_box_criteria over DEFAULT_BUY_BOX_CRITERIA —
        same reasoning as Organization.get_debt_assumptions(): a
        workspace that's never touched buy-box settings has
        buy_box_criteria == {} (the JSONField default), and a
        partially-filled-in dict should still fall back to the defaults
        (importantly, enabled=False) for every key it hasn't set.
        """
        merged = dict(self.DEFAULT_BUY_BOX_CRITERIA)
        merged.update(self.buy_box_criteria or {})
        return merged

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['organization', 'name'], name='unique_workspace_name_per_org'),
        ]
        ordering = ['name']

    def __str__(self):
        return f"{self.name} ({self.organization.name})"

    def _generate_unique_inbound_slug(self) -> str:
        """
        Base slug from org id + workspace name (human-recognizable in the
        resulting email address), suffixed with a short random hex string
        to guarantee global uniqueness without a second DB round-trip in
        the common case. The loop below is belt-and-suspenders only — a
        6-hex-char collision on top of an already-distinguishing base is
        astronomically unlikely, but this mirrors the rest of this
        codebase's "never assume, verify" discipline rather than trusting
        that math.
        """
        base = slugify(f"{self.organization_id}-{self.name}")[:40] or 'workspace'
        for _ in range(5):
            candidate = f"{base}-{uuid.uuid4().hex[:6]}"
            if not TeamWorkspace.objects.filter(inbound_email_slug=candidate).exists():
                return candidate
        # Fallback if every attempt above somehow collided: a pure UUID4
        # hex string has no meaningful collision probability at all.
        return uuid.uuid4().hex

    def save(self, *args, **kwargs):
        if not self.inbound_email_slug:
            self.inbound_email_slug = self._generate_unique_inbound_slug()
        super().save(*args, **kwargs)

    @property
    def inbound_email_address(self) -> str:
        """
        The full address shown in Workspace Settings and matched by
        InboundEmailWebhookView — e.g.
        'deals-acme-dispositions-a1b2c3@inboundunderwriting.com'. Built
        from settings.INBOUND_EMAIL_DOMAIN rather than stored, so changing
        that setting (e.g. switching domains) instantly updates every
        workspace's displayed address without a data migration.
        """
        return f"deals-{self.inbound_email_slug}@{django_settings.INBOUND_EMAIL_DOMAIN}"


class CustomUserManager(BaseUserManager):
    def create_user(self, email, password=None, **extra_fields):
        if not email:
            raise ValueError('The Email field must be set')
        email = self.normalize_email(email)
        user = self.model(email=email, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, email, password=None, **extra_fields):
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        extra_fields.setdefault('is_admin', True)
        return self.create_user(email, password, **extra_fields)

class CustomUser(AbstractBaseUser, PermissionsMixin):
    ROLE_CHOICES = [
        ('admin', 'Company Admin'), # e.g., The CEO of the company
        ('member', 'Team Member'),  # e.g., The employees
    ]

    email = models.EmailField(unique=True, primary_key=True)
    first_name = models.CharField(max_length=100)
    last_name = models.CharField(max_length=100)
    
    organization = models.ForeignKey(Organization, on_delete=models.CASCADE, null=True, blank=True, related_name='users')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='member')

    # Feature 3: nullable so every existing account is completely unaffected
    # until an org admin explicitly assigns them to a workspace. Drives the
    # Deal Pipeline visibility rule (see ai_service/permissions.py): an org
    # admin (role='admin') always sees every workspace in their organization
    # regardless of this field; within a workspace, visibility is further
    # split by workspace_role below (Feature 4); a member with no workspace
    # assigned yet (the expected state during rollout) falls back to seeing
    # only the projects they personally created, rather than either their
    # whole org or nothing.
    # Application layer (not the database — see TeamWorkspace's docstring)
    # enforces that this always belongs to the same organization as the
    # user themselves.
    workspace = models.ForeignKey(
        'TeamWorkspace', on_delete=models.SET_NULL, null=True, blank=True, related_name='members'
    )

    WORKSPACE_ROLE_CHOICES = [
        ('analyst', 'Analyst'),             # The analyst seat — own-created deals only.
        ('head', 'Head of Underwriting'),   # The "Master Seat" — sees every deal in
                                             # this workspace, plus team usage metrics
                                             # and every override audit log on the team.
    ]

    # Feature 4 ("Head of Underwriting" hierarchy): a SECOND, workspace-scoped
    # permission axis layered on top of `role` above. `role` is org-wide
    # (billing/admin); this only matters once a user actually has a
    # workspace, and is simply ignored otherwise — see visible_projects_for
    # in ai_service/permissions.py for the full precedence (org admin >
    # workspace head > plain member/no-workspace, both of which fall back to
    # own-created-only). Default 'analyst' is deliberately the safer choice
    # for every NEW workspace assignment going forward — but see this
    # field's migration for why every user who ALREADY had a workspace
    # before this field existed is backfilled to 'head' instead, so this
    # rollout doesn't silently shrink anyone's existing visibility.
    workspace_role = models.CharField(max_length=10, choices=WORKSPACE_ROLE_CHOICES, default='analyst')

    is_active = models.BooleanField(default=True)
    is_staff = models.BooleanField(default=False)
    is_admin = models.BooleanField(default=False) # ONLY FOR YOU


    date_joined = models.DateTimeField(auto_now_add=True)

    last_active = models.DateTimeField(null=True, blank=True)

    objects = CustomUserManager()

    USERNAME_FIELD = 'email'
    REQUIRED_FIELDS = ['first_name', 'last_name']

    def save(self, *args, **kwargs):
        # STRICT ENFORCEMENT: Only your specific email gets platform admin rights.
        # Everyone else is forcefully denied, preventing any UI or database accidents.
        if self.email.lower() == 'renaldovanreden@gmail.com':
            self.is_admin = True
            self.is_staff = True
            self.is_superuser = True
            self.is_active = True
        else:
            self.is_admin = False
            self.is_staff = False
            self.is_superuser = False
            
        super().save(*args, **kwargs)

    def __str__(self):
        org_name = self.organization.name if self.organization else 'System Admin'
        return f"{self.email} ({org_name})"

# backend activity logging model

class ActivityLog(models.Model):
    ACTION_CHOICES = [
        ('LOGIN', 'User Logged In'),
        ('LOGOUT', 'User Logged Out'),
        ('CREATE_ACCOUNT', 'Account Created'),
        ('DELETE_ACCOUNT', 'Account Deleted'),
        ('TOGGLE_ACCOUNT', 'Account Status Changed'),
        ('UPDATE_PLAN', 'Account Plan Updated'), #account plan updated
        ('CREATE_USER', 'Team Member Added'),
        ('DELETE_USER', 'Team Member Removed'),
        ('TOGGLE_USER', 'Team Member Status Changed'),
    ]

    user = models.ForeignKey(CustomUser, on_delete=models.SET_NULL, null=True, blank=True, related_name='activity_logs')
    action = models.CharField(max_length=50, choices=ACTION_CHOICES)
    target = models.CharField(max_length=255, help_text="The company or user affected (e.g., 'IMPI Engineering')")
    details = models.TextField(blank=True, null=True, help_text="Extra context (e.g., 'Changed status to Inactive')")
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-timestamp'] # Newest first

    def __str__(self):
        return f"{self.action} - {self.target} by {self.user.email if self.user else 'System'}"

class SystemMessage(models.Model):
    """
    Admin-posted announcements shown to every user (e.g. maintenance
    windows, upcoming changes). Displayed in the dashboard sidebar, per the
    design decision — not a dismissible top banner.
 
    Lifecycle: auto-expires 30 days after posting (expires_at is set once,
    at creation, and never changes), but an admin can also take a message
    down early via is_active — e.g. once the maintenance window has already
    passed. A message is only shown to users when BOTH is_active=True AND
    expires_at is still in the future.
    """
    MESSAGE_TYPE_CHOICES = [
        ('info', 'Info'),
        ('warning', 'Warning'),
        ('critical', 'Critical'),
    ]
 
    MESSAGE_LIFETIME_DAYS = 30
 
    message = models.TextField()
    message_type = models.CharField(max_length=20, choices=MESSAGE_TYPE_CHOICES, default='info')
 
    created_by = models.ForeignKey(
        CustomUser,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='posted_messages'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    expires_at = models.DateTimeField(editable=False)
 
    is_active = models.BooleanField(default=True)  # lets an admin remove it before natural expiry
 
    class Meta:
        ordering = ['-created_at']
 
    def save(self, *args, **kwargs):
        if not self.expires_at:
            self.expires_at = timezone.now() + timedelta(days=self.MESSAGE_LIFETIME_DAYS)
        super().save(*args, **kwargs)
 
    def __str__(self):
        return f"[{self.message_type}] {self.message[:50]} (posted {self.created_at.date()})"

class IssueReport(models.Model):
    """
    User-submitted issue reports, reviewed by the admin on a dedicated
    /admin/issues page. Mirror image of SystemMessage (that's admin -> users,
    this is users -> admin).
 
    v1 scope: submit + admin views + admin marks resolved. No in-app reply/
    threading — if the admin needs to respond, that happens outside the
    platform (email, call, etc.), same as trial account credential handoff.
    """
    CATEGORY_CHOICES = [
        ('bug', 'Bug'),
        ('question', 'Question'),
        ('feature_request', 'Feature Request'),
    ]
 
    STATUS_CHOICES = [
        ('open', 'Open'),
        ('resolved', 'Resolved'),
    ]
 
    reported_by = models.ForeignKey(
        CustomUser,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='issue_reports'
    )
    # Denormalized (not just derived via reported_by.organization) so the
    # report still shows which org it came from even if the user is later
    # deleted — same reasoning as ProcessedDocument/AnalysisReport.
    organization = models.ForeignKey(
        Organization,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='issue_reports'
    )
 
    category = models.CharField(max_length=20, choices=CATEGORY_CHOICES, default='bug')
    description = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='open')
 
    created_at = models.DateTimeField(auto_now_add=True)
    resolved_at = models.DateTimeField(null=True, blank=True)
 
    class Meta:
        ordering = ['-created_at']
 
    def __str__(self):
        return f"[{self.category}] {self.description[:50]} ({self.status})"