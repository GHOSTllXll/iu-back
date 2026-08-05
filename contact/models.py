from django.db import models

class Contact(models.Model):
    company = models.CharField(max_length=255)
    company_size = models.CharField(max_length=100)
    contact_name = models.CharField(max_length=255)
    email = models.EmailField()
    phone = models.CharField(max_length=50)
    message = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)

    STATUS_CHOICES = (
        ("new", "New"),
        ("contacted", "Contacted"),
        ("closed", "Closed"),
    )

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="new"
    )

    def __str__(self):
        return self.company