# users/models.py
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models
from django.utils import timezone
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

    objects = OrganizationManager()

    def __str__(self):
        return self.name

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