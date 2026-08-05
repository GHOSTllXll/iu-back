# users/models.py
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models

class OrganizationManager(models.Manager):
    pass

class Organization(models.Model):
    PLAN_CHOICES = [
        ('basic', 'Basic'),
        ('professional', 'Professional'),
        ('enterprise', 'Enterprise'),
    ]
    STATUS_CHOICES = [
        ('active', 'Active'),
        ('pending', 'Pending'),
        ('inactive', 'Inactive'),
    ]

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