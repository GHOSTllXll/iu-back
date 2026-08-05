# backend/users/middleware.py
from django.utils import timezone
from django.contrib.auth.models import AnonymousUser

class ActiveUserMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        
        # Only update for authenticated users (ignore anonymous/public requests)
        if hasattr(request, 'user') and not request.user.is_anonymous:
            # Update the last_active field
            request.user.last_active = timezone.now()
            request.user.save(update_fields=['last_active'])
            
        return response