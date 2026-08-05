# backend/users/serializers.py
from rest_framework import serializers
from .models import CustomUser, Organization, ActivityLog

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
    
    class Meta:
        model = CustomUser
        fields = ['email', 'first_name', 'last_name', 'role', 'is_admin', 'is_active', 'date_joined', 'organization_name']

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

        # Create the user (password is automatically hashed by the model's manager)
        user = CustomUser.objects.create_user(
            password=password,
            organization=org,
            **validated_data
        )
        return user

class ActivityLogSerializer(serializers.ModelSerializer):
    # Added allow_null=True here
    user_email = serializers.CharField(source='user.email', read_only=True, allow_null=True)
    action_display = serializers.CharField(source='get_action_display', read_only=True)

    class Meta:
        model = ActivityLog
        fields = ['id', 'action', 'action_display', 'target', 'details', 'user_email', 'timestamp']