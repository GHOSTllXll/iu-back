from django.contrib import admin
from .models import Contact


@admin.register(Contact)
class ContactAdmin(admin.ModelAdmin):

    list_display = (
        "company",
        "contact_name",
        "email",
        "status",
        "created_at",
    )

    search_fields = (
        "company",
        "contact_name",
        "email",
    )

    list_filter = (
        "status",
        "created_at",
    )