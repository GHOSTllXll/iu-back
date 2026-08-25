# backend/users/middleware.py
from django.utils import timezone
from django.contrib.auth.models import AnonymousUser
from django.db import DatabaseError, InterfaceError

class ActiveUserMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response
    def __call__(self, request):
        response = self.get_response(request)
        
        if hasattr(request, 'user') and not request.user.is_anonymous:
            try:
                request.user.last_active = timezone.now()
                request.user.save(update_fields=['last_active'])
            except (DatabaseError, InterfaceError):
                # Non-critical housekeeping field. If the DB connection went
                # stale during a long-running request (e.g. a slow AI call
                # exceeding MySQL's idle connection timeout), don't let this
                # trivial update crash an otherwise-successful response.
                pass
            
        return response

class NoCacheAPIMiddleware:
    """
    Prevents browsers from caching ANY /api/ response. Without this, a
    browser can serve a stale cached GET response even after a mutating
    action (create/update/delete) elsewhere — exactly the bug we found and
    fixed for the login/user endpoint. This applies the same fix globally,
    to every API endpoint, instead of adding Cache-Control headers to each
    view individually.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        if request.path.startswith('/api/'):
            response['Cache-Control'] = 'no-store, no-cache, must-revalidate, private'
            response['Pragma'] = 'no-cache'
        return response