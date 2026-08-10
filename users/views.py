from django.shortcuts import render

# users/views.py
import urllib.parse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.permissions import IsAuthenticated
from django.contrib.auth import authenticate
from rest_framework_simplejwt.tokens import RefreshToken
from django.utils import timezone
from datetime import timedelta

from .utils import log_activity
from .models import CustomUser, Organization, ActivityLog
from .serializers import LoginSerializer, UserSerializer, AdminCreateUserSerializer, OrganizationSerializer, ActivityLogSerializer

from django.contrib.auth.tokens import PasswordResetTokenGenerator
from django.utils.encoding import force_bytes, force_str
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.core.mail import send_mail
from django.conf import settings

# Add this near the top level (module scope), alongside the other module-level
# objects — reused across both views below.
password_reset_token_generator = PasswordResetTokenGenerator()

# Replace the existing LoginView class in users/views.py with this version.
# Only the authentication logic changed — cookie-setting and activity
# logging at the bottom are untouched.

class LoginView(APIView):
    def post(self, request):
        serializer = LoginSerializer(data=request.data)
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        email = serializer.validated_data['email']
        password = serializer.validated_data['password']

        # NOTE: deliberately NOT using Django's authenticate() here. It calls
        # ModelBackend.authenticate() under the hood, which silently refuses
        # to authenticate an inactive user — it returns None for them, making
        # a deactivated account indistinguishable from a wrong password. That
        # meant the "account is inactive" branch below could never actually
        # fire. Looking the user up manually and checking the password
        # ourselves lets us tell the two cases apart and return the right
        # message for each.
        try:
            user = CustomUser.objects.get(email=email)
        except CustomUser.DoesNotExist:
            return Response({'detail': 'Invalid email or password.'}, status=status.HTTP_401_UNAUTHORIZED)

        if not user.check_password(password):
            return Response({'detail': 'Invalid email or password.'}, status=status.HTTP_401_UNAUTHORIZED)

        if not user.is_active:
            return Response(
                {'detail': 'Your account has been deactivated. Please contact support.'},
                status=status.HTTP_403_FORBIDDEN
            )

        # Generate JWT tokens
        refresh = RefreshToken.for_user(user)

        # Prepare the response
        response = Response({
            'detail': 'Login successful',
            'user': UserSerializer(user).data
        })

        # SECURITY: Set tokens in HttpOnly cookies
        # secure=False is for local dev. Change to True in production (requires HTTPS).
        cookie_kwargs = {
            'httponly': True,
            'secure': False,
            'samesite': 'Lax',
            'path': '/'
        }

        response.set_cookie(key='access_token', value=str(refresh.access_token), **cookie_kwargs)
        response.set_cookie(key='refresh_token', value=str(refresh), **cookie_kwargs)

        log_activity(
            user=user,
            action='LOGIN',
            target=user.email,
            ip_address=request.META.get('REMOTE_ADDR')
        )

        return response

class LogoutView(APIView):
    permission_classes = [IsAuthenticated]
    
    def post(self, request):
        response = Response({'detail': 'Logged out successfully'})
        # Delete the cookies
        response.delete_cookie('access_token', path='/')
        response.delete_cookie('refresh_token', path='/')

        log_activity(
            user=request.user, 
            action='LOGOUT', 
            target=request.user.email,
            ip_address=request.META.get('REMOTE_ADDR')
        )

        return response

class CurrentUserView(APIView):
    permission_classes = [IsAuthenticated]
    
    def get(self, request):
        # request.user is populated by our custom CookieJWTAuthentication
        return Response(UserSerializer(request.user).data)

class AdminCreateUserView(APIView):
    permission_classes = [IsAuthenticated]
    
    def post(self, request):
        # STRICT ADMIN CHECK: Only YOU can access this endpoint.
        if not request.user.is_admin:
            return Response({'detail': 'You do not have permission to create users.'}, status=status.HTTP_403_FORBIDDEN)
            
        serializer = AdminCreateUserSerializer(data=request.data)
        if serializer.is_valid():
            user = serializer.save()
            log_activity(
            user=request.user,
            action='CREATE_USER',
            target=user.email,
            details=f"Added to {request.data.get('company_name', 'Unknown Company')}",
            ip_address=request.META.get('REMOTE_ADDR')
        )

            response_data = {
                'detail': f'User {user.email} created successfully.',
                'user': UserSerializer(user).data,
                'org_was_created': serializer.org_was_created,
            }

            # SAFETY NET: only warn when (a) the org already existed, AND
            # (b) a subscription_plan was ACTUALLY present in the raw request
            # (not just DRF's field default kicking in silently), AND (c) it
            # doesn't match the org's real current plan. This keeps the
            # warning from firing on account/[id].vue's "Add Team Member"
            # flow, which never offers a plan choice and no longer sends the
            # field at all.
            plan_was_explicitly_sent = 'subscription_plan' in request.data
            if (not serializer.org_was_created
                    and plan_was_explicitly_sent
                    and serializer.requested_plan != serializer.org_actual_plan):
                response_data['warning'] = (
                    f"'{request.data.get('company_name')}' already exists — the plan you "
                    f"selected was NOT applied. This organization's current plan is "
                    f"'{serializer.org_actual_plan}'. To change it, use the plan dropdown "
                    f"on that organization's account page."
                )
                response_data['org_actual_plan'] = serializer.org_actual_plan

            return Response(response_data, status=status.HTTP_201_CREATED)
            
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
    
class AdminOrganizationsView(APIView):
    permission_classes = [IsAuthenticated]
    
    def get(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        # Fetch all organizations, newest first
        orgs = Organization.objects.all().order_by('-date_created')
        serializer = OrganizationSerializer(orgs, many=True)
        return Response(serializer.data)

class AdminOrganizationUsersView(APIView):
    permission_classes = [IsAuthenticated]
    
    def get(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)
            
        # Get all users linked to this organization
        users = CustomUser.objects.filter(organization=org).order_by('-date_joined')
        serializer = UserSerializer(users, many=True)
        return Response(serializer.data)

class AdminToggleUserStatusView(APIView):
    permission_classes = [IsAuthenticated]
    
    def post(self, request, email):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            # URL decode the email (e.g., turns 'user%40test.com' back to 'user@test.com')
            decoded_email = urllib.parse.unquote(email)
            user = CustomUser.objects.get(email=decoded_email)
            
            # SECURITY: Prevent deactivating yourself (the system admin)
            if user.is_admin:
                return Response({'detail': 'Cannot deactivate the system administrator.'}, status=status.HTTP_400_BAD_REQUEST)
                
            # Toggle the status
            user.is_active = not user.is_active
            user.save()

            log_activity(
                user=request.user,
                action='TOGGLE_USER',
                target=user.email,
                details=f"Status changed to {'Active' if user.is_active else 'Inactive'}",
                ip_address=request.META.get('REMOTE_ADDR')
            )
            
            return Response({'detail': f'User status updated to {"Active" if user.is_active else "Inactive"}'})
            
        except CustomUser.DoesNotExist:
            return Response({'detail': 'User not found'}, status=status.HTTP_404_NOT_FOUND)
        

# added views for admin to toggle organization status and delete organization
class AdminToggleOrganizationStatusView(APIView):
    permission_classes = [IsAuthenticated]
    
    def patch(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)
            
        # Toggle the status
        new_status = 'inactive' if org.account_status == 'active' else 'active'
        org.account_status = new_status
        org.save()

        log_activity(
            user=request.user,
            action='TOGGLE_ACCOUNT',
            target=org.name,
            details=f"Status changed to {new_status}",
            ip_address=request.META.get('REMOTE_ADDR')
        )
        
        # CASCADING LOGIC: Update all users under this organization
        if new_status == 'inactive':
            # Deactivate all users in this company
            CustomUser.objects.filter(organization=org).update(is_active=False)
        else:
            # Optional: Reactivate them when the company is reactivated
            CustomUser.objects.filter(organization=org).update(is_active=True)
            
        serializer = OrganizationSerializer(org)
        return Response(serializer.data)


class AdminDeleteOrganizationView(APIView):
    permission_classes = [IsAuthenticated]
    
    def delete(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            org = Organization.objects.get(id=org_id)
            # CASCADE DELETE: Because of on_delete=models.CASCADE in your CustomUser model,
            # deleting the organization will automatically delete all linked users safely.

            log_activity(
                user=request.user,
                action='DELETE_ACCOUNT',
                target=org.name,
                details="Permanently deleted organization and all users",
                ip_address=request.META.get('REMOTE_ADDR')
            )
            org.delete()
            return Response({'detail': 'Organization deleted successfully'}, status=status.HTTP_204_NO_CONTENT)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)
        
# NEW: Admin Delete User View
class AdminDeleteUserView(APIView):
    permission_classes = [IsAuthenticated]
    
    def delete(self, request, email):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            # URL decode the email (e.g., turns 'user%40test.com' back to 'user@test.com')
            decoded_email = urllib.parse.unquote(email)
            user = CustomUser.objects.get(email=decoded_email)
            
            # SECURITY: Prevent deleting the system administrator
            if user.is_admin:
                return Response({'detail': 'Cannot delete the system administrator.'}, status=status.HTTP_400_BAD_REQUEST)
                
            user_email = user.email  # Store email for logging before deletion
            # Permanently delete the user
            user.delete()
            log_activity(
                user=request.user,
                action='DELETE_USER',
                target=user_email,
                details="Permanently deleted user",
                ip_address=request.META.get('REMOTE_ADDR')
            )
            return Response({'detail': 'User deleted successfully'}, status=status.HTTP_204_NO_CONTENT)
            
        except CustomUser.DoesNotExist:
            return Response({'detail': 'User not found'}, status=status.HTTP_404_NOT_FOUND)
        
class AdminActiveUsersCountView(APIView):
    permission_classes = [IsAuthenticated]
    
    def get(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        # Define "currently logged in" as active within the last 15 minutes
        fifteen_minutes_ago = timezone.now() - timedelta(minutes=15)
        
        # Count users who have a last_active timestamp newer than 15 mins ago
        active_now_count = CustomUser.objects.filter(
            last_active__gte=fifteen_minutes_ago,
            is_active=True
        ).count()
        
        return Response({'active_now': active_now_count})

class AdminActivityLogView(APIView):
    permission_classes = [IsAuthenticated]
    
    def get(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        # Fetch the last 50 logs, newest first
        logs = ActivityLog.objects.select_related('user').all()[:50]
        serializer = ActivityLogSerializer(logs, many=True)
        
        return Response(serializer.data)
    
class AdminUpdateOrganizationPlanView(APIView):
    permission_classes = [IsAuthenticated]
    
    def patch(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
            
        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)
            
        new_plan = request.data.get('subscription_plan')
        valid_plans = ['basic', 'professional', 'enterprise']
        
        if new_plan not in valid_plans:
            return Response({'detail': 'Invalid plan selected.'}, status=status.HTTP_400_BAD_REQUEST)
            
        old_plan_display = org.get_subscription_plan_display()
        
        # Update the plan
        org.subscription_plan = new_plan
        org.save()
        
        # Log the activity
        log_activity(
            user=request.user,
            action='UPDATE_PLAN',
            target=org.name,
            details=f"Plan changed from {old_plan_display} to {org.get_subscription_plan_display()}",
            ip_address=request.META.get('REMOTE_ADDR')
        )
        
        serializer = OrganizationSerializer(org)
        return Response(serializer.data)

class PasswordResetRequestView(APIView):
    """
    Public endpoint (no auth required) — the "Forgot password?" flow's first
    step. Takes an email, and IF a matching active user exists, emails them a
    reset link containing a signed token.

    SECURITY: always returns the same generic success message regardless of
    whether the email actually exists in the system. This is deliberate — it
    stops someone from using this endpoint to check which email addresses
    are registered on the platform.
    """
    def post(self, request):
        email = request.data.get('email', '').strip()

        if not email:
            return Response({'detail': 'Email is required.'}, status=status.HTTP_400_BAD_REQUEST)

        generic_response = Response({
            'detail': 'If an account exists with that email, a password reset link has been sent.'
        })

        try:
            user = CustomUser.objects.get(email=email)
        except CustomUser.DoesNotExist:
            return generic_response  # don't reveal whether the email exists

        if not user.is_active:
            # Deliberately still returns the generic message — same reasoning
            # as above, don't leak account status through this endpoint either.
            return generic_response

        uid = urlsafe_base64_encode(force_bytes(user.pk))
        token = password_reset_token_generator.make_token(user)
        reset_link = f"{settings.FRONTEND_BASE_URL}/reset-password?uid={uid}&token={token}"

        # NOTE: requires EMAIL_BACKEND to be configured in settings.py. For
        # local dev, Django's console backend prints the email to your
        # terminal instead of actually sending it — zero config needed:
        #     EMAIL_BACKEND = 'django.core.mail.backends.console.EmailBackend'
        # For production, you'll need a real provider (SendGrid, AWS SES,
        # Mailgun, etc.) with SMTP credentials in settings.py / .env.
        send_mail(
            subject="Reset your Inbound Underwriting password",
            message=(
                f"Hi {user.first_name},\n\n"
                f"Click the link below to reset your password:\n\n"
                f"{reset_link}\n\n"
                f"This link expires in {settings.PASSWORD_RESET_TIMEOUT // 3600 if hasattr(settings, 'PASSWORD_RESET_TIMEOUT') else 24} hour(s). "
                f"If you didn't request this, you can safely ignore this email."
            ),
            from_email=getattr(settings, 'DEFAULT_FROM_EMAIL', 'noreply@inboundunderwriting.com'),
            recipient_list=[user.email],
            fail_silently=False,
        )

        return generic_response


class PasswordResetConfirmView(APIView):
    """
    Public endpoint (no auth required) — the "Forgot password?" flow's second
    step. Takes the uid/token from the emailed link plus a new password, and
    actually sets it.
    """
    def post(self, request):
        uid = request.data.get('uid', '')
        token = request.data.get('token', '')
        new_password = request.data.get('new_password', '')

        if not all([uid, token, new_password]):
            return Response(
                {'detail': 'uid, token, and new_password are all required.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        if len(new_password) < 8:
            return Response(
                {'detail': 'Password must be at least 8 characters long.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        try:
            user_pk = force_str(urlsafe_base64_decode(uid))
            user = CustomUser.objects.get(pk=user_pk)
        except (TypeError, ValueError, OverflowError, CustomUser.DoesNotExist):
            return Response({'detail': 'Invalid reset link.'}, status=status.HTTP_400_BAD_REQUEST)

        if not password_reset_token_generator.check_token(user, token):
            # Covers both a wrong/tampered token AND an expired one (the
            # token generator encodes a timestamp check internally, governed
            # by settings.PASSWORD_RESET_TIMEOUT, default 3 days).
            return Response(
                {'detail': 'This reset link is invalid or has expired. Please request a new one.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        user.set_password(new_password)
        user.save()

        log_activity(
            user=user,
            action='LOGIN',  # no dedicated action type exists for this yet — reuses closest existing one
            target=user.email,
            details="Password reset via forgot-password flow",
            ip_address=request.META.get('REMOTE_ADDR')
        )

        return Response({'detail': 'Password reset successfully. You can now log in with your new password.'})