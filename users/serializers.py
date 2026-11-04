# backend/users/serializers.py
from rest_framework import serializers
from .models import CustomUser, Organization, ActivityLog, TeamWorkspace

#  NEW: Serializer for the Accounts Dashboard
class OrganizationSerializer(serializers.ModelSerializer):
    user_count = serializers.SerializerMethodField()
    inactive_user_count = serializers.SerializerMethodField()
    plan_display = serializers.CharField(source='get_subscription_plan_display') # e.g., "Pro"
    status_display = serializers.CharField(source='get_account_status_display')  # e.g., "Active"
    created = serializers.SerializerMethodField()
    contact_full = serializers.SerializerMethodField()

    class Meta:
        model = Organization
        fields = [
            'id', 'name', 'primary_email', 'contact_full', 
            'plan_display', 'status_display', 'user_count', 'inactive_user_count', 'created'
        ]

    def get_user_count(self, obj):
        return obj.users.filter(is_active=True).count() # Count only active users
    
    def get_inactive_user_count(self, obj):
        return obj.users.filter(is_active=False).count() # Count only inactive users

    def get_created(self, obj):
        return obj.date_created.strftime("%d %b %Y") # Formats as "18 Jun 2026"

    def get_contact_full(self, obj):
        return f"{obj.primary_contact_name} {obj.primary_contact_surname}"

class UserSerializer(serializers.ModelSerializer):
    # allow_null=True prevents crashes when the user has no organization (like you, the admin!)
    organization_name = serializers.CharField(source='organization.name', read_only=True, allow_null=True)

    # Drives the Institutional PPM Batch Console's UI gating on the frontend
    # (sidebar link visibility + the page's own access gate) — a
    # SerializerMethodField rather than a dotted `source=` lookup so a
    # user with no organization at all returns False instead of blowing up.
    has_ppm_access = serializers.SerializerMethodField()

    # Feature 4: drives the Team Workspaces section's "View Usage" link
    # gating on settings.vue — a workspace Head needs to know THEIR OWN
    # workspace_id/workspace_role to tell their own workspace apart from
    # every other one in the list (never trusted for the actual permission
    # check itself, which WorkspaceUsageView re-derives server-side from
    # request.user). Declared explicitly rather than left to ModelSerializer
    # auto-detection — 'workspace_id' is the FK's attname, not its field
    # name ('workspace'), so it isn't one of the fields ModelSerializer can
    # build automatically; allow_null=True for the no-workspace case.
    workspace_id = serializers.IntegerField(read_only=True, allow_null=True)

    class Meta:
        model = CustomUser
        fields = [
            'email', 'first_name', 'last_name', 'role', 'is_admin', 'is_active', 'date_joined',
            'organization_name', 'has_ppm_access', 'workspace_id', 'workspace_role',
        ]

    def get_has_ppm_access(self, obj):
        return bool(obj.organization and obj.organization.has_ppm_access)

class LoginSerializer(serializers.Serializer):
    email = serializers.EmailField()
    password = serializers.CharField(write_only=True)

class AdminCreateUserSerializer(serializers.ModelSerializer):
    password = serializers.CharField(write_only=True, required=True)
    
    #  CHANGED: Match the frontend's 'company_name' exactly
    company_name = serializers.CharField(required=True, write_only=True) 
    subscription_plan = serializers.ChoiceField(choices=Organization.PLAN_CHOICES, default='basic', write_only=True)

    class Meta:
        model = CustomUser
        #  CHANGED: Use 'company_name' here to match the frontend payload
        fields = ['email', 'first_name', 'last_name', 'password', 'role', 'company_name', 'subscription_plan']

    def create(self, validated_data):
        # Pop the custom fields from the validated data
        company_name = validated_data.pop('company_name').strip()
        password = validated_data.pop('password')
        plan = validated_data.pop('subscription_plan', 'basic')
        
        # Get or Create the Organization
        org, created = Organization.objects.get_or_create(
            name=company_name,
            defaults={
                'primary_contact_name': validated_data['first_name'],
                'primary_contact_surname': validated_data['last_name'],
                'primary_email': validated_data['email'],
                'subscription_plan': plan,
                'team_size_limit': 10, # Default limit, you can change later via Django Admin
                'account_status': 'active',
                'is_paid': False
            }
        )

        # SAFETY NET: get_or_create's `defaults` only applies when a NEW org is
        # being created. If `company_name` matched an EXISTING organization,
        # whatever plan was selected on this form is silently ignored — the
        # org keeps its current plan. Stash what actually happened here so
        # the view can warn the admin instead of that being invisible.
        self.org_was_created = created
        self.org_actual_plan = org.subscription_plan
        self.requested_plan = plan

        # Create the user (password is automatically hashed by the model's manager)
        user = CustomUser.objects.create_user(
            password=password,
            organization=org,
            **validated_data
        )
        return user

class AssignUserWorkspaceSerializer(serializers.Serializer):
    """
    Feature 3: validates a workspace assignment BEFORE it's applied — the
    application-layer enforcement (praised in review as "the brilliant
    application-layer constraint") that a CustomUser's workspace always
    belongs to that same user's organization. The database can't express
    this itself (see TeamWorkspace's own docstring for why — it's a
    constraint across two different FKs). workspace_id=None is valid and
    means "unassign" — the member then falls back to the own-created-
    projects-only visibility rule (see ai_service/permissions.py).

    Requires 'organization' in the serializer context — the CALLER's own
    organization, resolved by the view from request.user.organization (the
    self-serve endpoint) or the org_id in the URL (the platform-admin
    mirror) — never trusted from the request body itself.

    Feature 4: also accepts an optional 'workspace_role' ('analyst' or
    'head') so a single PATCH can set a member's workspace assignment and
    their seat within it together. Omitting it leaves the member's current
    workspace_role untouched (the view only applies a key that was actually
    present in the request - see MyOrganizationMemberWorkspaceView.patch /
    AdminOrganizationUserWorkspaceView.patch). This field is intentionally
    NOT validated against workspace_id being non-null here - a head with no
    workspace assigned is harmless (permissions.py's visible_projects_for
    requires BOTH workspace_id is not None AND workspace_role == 'head'
    before granting workspace-wide visibility), so the view is free to
    accept workspace_role on its own, e.g. to pre-promote someone ahead of
    assigning their workspace.
    """
    workspace_id = serializers.IntegerField(required=False, allow_null=True)
    workspace_role = serializers.ChoiceField(
        choices=CustomUser.WORKSPACE_ROLE_CHOICES, required=False
    )

    def validate_workspace_id(self, value):
        if value is None:
            return value
        organization = self.context['organization']
        if not TeamWorkspace.objects.filter(pk=value, organization=organization).exists():
            raise serializers.ValidationError(
                "This workspace does not belong to this user's organization."
            )
        return value


class ActivityLogSerializer(serializers.ModelSerializer):
    # Added allow_null=True here
    user_email = serializers.CharField(source='user.email', read_only=True, allow_null=True)
    action_display = serializers.CharField(source='get_action_display', read_only=True)

    class Meta:
        model = ActivityLog
        fields = ['id', 'action', 'action_display', 'target', 'details', 'user_email', 'timestamp']