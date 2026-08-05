# users/authentication.py
from rest_framework_simplejwt.authentication import JWTAuthentication

class CookieJWTAuthentication(JWTAuthentication):
    """
    Custom authentication class that reads the JWT from an HttpOnly cookie
    instead of the Authorization header. Prevents XSS attacks from accessing the token via JavaScript.
    """
    def authenticate(self, request):
        # Look for the token in the 'access_token' cookie
        raw_token = request.COOKIES.get('access_token')
        
        if raw_token is None:
            return None
            
        try:
            validated_token = self.get_validated_token(raw_token)
        except Exception:
            return None
            
        return self.get_user(validated_token), validated_token