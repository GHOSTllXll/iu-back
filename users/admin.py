# users/admin.py
from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import Organization, CustomUser

@admin.register(Organization)
class OrganizationAdmin(admin.ModelAdmin):
    list_display = ('name', 'primary_email', 'subscription_plan', 'team_size_limit', 'is_paid', 'account_status')
    list_filter = ('account_status', 'subscription_plan', 'is_paid')
    search_fields = ('name', 'primary_email')

class CustomUserAdmin(UserAdmin):
    model = CustomUser
    list_display = ('email', 'first_name', 'last_name', 'organization', 'role', 'is_active', 'is_admin')
    list_filter = ('is_active', 'role', 'is_admin', 'organization')
    
    # Notice 'is_admin' is NOT in the fieldsets below. It cannot be edited via the UI.
    fieldsets = (
        (None, {'fields': ('email', 'password')}),
        ('Personal Info', {'fields': ('first_name', 'last_name')}),
        ('Organization & Role', {'fields': ('organization', 'role')}),
        ('Permissions', {'fields': ('is_active', 'is_staff', 'is_superuser')}),
    )
    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('email', 'first_name', 'last_name', 'organization', 'role', 'password1', 'password2'),
        }),
    )
    search_fields = ('email', 'first_name', 'last_name')
    ordering = ('email',)

admin.site.register(CustomUser, CustomUserAdmin)