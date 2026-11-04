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
from .models import CustomUser, Organization, ActivityLog, SystemMessage, IssueReport, TeamWorkspace
from .serializers import (
    LoginSerializer, UserSerializer, AdminCreateUserSerializer, OrganizationSerializer,
    ActivityLogSerializer, AssignUserWorkspaceSerializer,
)

from django.contrib.auth.tokens import PasswordResetTokenGenerator
from django.utils.encoding import force_bytes, force_str
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.core.mail import send_mail
from django.conf import settings

from rest_framework.throttling import AnonRateThrottle, SimpleRateThrottle

import secrets
import string
import io
import openpyxl
from django.core.files.base import ContentFile

from ai_service.models import ClientExcelTemplate, PropertyProject, AnalysisReport
from ai_service.custom_template_generator import (
    validate_template,
    TemplateValidationError,
    SCALAR_FIELD_CATALOG,
    ROW_FIELD_CATALOG,
)

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
            'secure': not settings.DEBUG,
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
    # Deliberately NOT IsAuthenticated. Logout's entire job is clearing
    # cookies — it shouldn't require a still-valid token to do that. If it
    # did, an expired-token session could never successfully log out at all
    # (the 401 from the auth check would block the request before it ever
    # reached this view), leaving the user stuck with dead cookies and no
    # way to clear them except manually.
    permission_classes = []
 
    def post(self, request):
        response = Response({'detail': 'Logged out successfully'})
        response.delete_cookie('access_token', path='/')
        response.delete_cookie('refresh_token', path='/')
 
        # Only log the activity if we actually have a valid, authenticated
        # user — with permission_classes=[] this endpoint can now be hit
        # with an expired/missing token too, in which case request.user is
        # AnonymousUser and there's nothing meaningful to log.
        if request.user and request.user.is_authenticated:
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
        response = Response(UserSerializer(request.user).data)
        response['Cache-Control'] = 'no-store, no-cache, must-revalidate, private'
        response['Pragma'] = 'no-cache'
        return response

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

class PasswordResetIPThrottle(AnonRateThrottle):
    """
    Caps how many reset REQUESTS a single IP can make, regardless of which
    email(s) they're targeting. Stops one source from hammering the endpoint.
    """
    scope = 'password_reset_ip'
 
 
class PasswordResetEmailThrottle(SimpleRateThrottle):
    """
    Caps how many reset requests can be made FOR A GIVEN EMAIL, regardless of
    which IP they come from. This is the layer that matters most here — an
    attacker rotating IPs (or using a botnet) could bypass IP-based throttling
    entirely, but flooding one specific person's inbox with reset emails
    still gets capped by this, since it's keyed on the target email itself.
    """
    scope = 'password_reset_email'
 
    def get_cache_key(self, request, view):
        email = request.data.get('email', '').strip().lower()
        if not email:
            # No email in the request — nothing to throttle on. The view's
            # own validation (email is required) handles this case; this
            # throttle just has nothing to key against.
            return None
        return self.cache_format % {
            'scope': self.scope,
            'ident': email,
        }
 
 
class PasswordResetConfirmThrottle(AnonRateThrottle):
    """
    Caps how many CONFIRM attempts (the uid/token/new_password step) a single
    IP can make. Tokens are cryptographically infeasible to brute-force, so
    this is defense-in-depth against automated retry spam, not a load-bearing
    security control on its own.
    """
    scope = 'password_reset_confirm_ip'

class PasswordResetRequestView(APIView):
    throttle_classes = [PasswordResetIPThrottle, PasswordResetEmailThrottle]
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
    throttle_classes = [PasswordResetConfirmThrottle]
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

class AdminCreateMessageView(APIView):
    """
    Admin posts a new announcement. Auto-expires 30 days from now
    (SystemMessage.save() sets expires_at automatically).
    URL: POST /api/admin/messages/
    """
    permission_classes = [IsAuthenticated]
 
    def post(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        message_text = request.data.get('message', '').strip()
        message_type = request.data.get('message_type', 'info')
 
        if not message_text:
            return Response({'detail': 'Message text is required.'}, status=status.HTTP_400_BAD_REQUEST)
 
        if message_type not in dict(SystemMessage.MESSAGE_TYPE_CHOICES):
            return Response({'detail': 'Invalid message type.'}, status=status.HTTP_400_BAD_REQUEST)
 
        msg = SystemMessage.objects.create(
            message=message_text,
            message_type=message_type,
            created_by=request.user,
        )
 
        return Response({
            'id': msg.id,
            'message': msg.message,
            'message_type': msg.message_type,
            'created_at': msg.created_at.isoformat(),
            'expires_at': msg.expires_at.isoformat(),
        }, status=status.HTTP_201_CREATED)
 
 
class AdminMessageListView(APIView):
    """
    Admin's own management view — ALL messages (active, inactive, expired),
    newest first, so the admin can see history and take early action.
    URL: GET /api/admin/messages/
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        now = timezone.now()
        messages = SystemMessage.objects.all()[:100]  # reasonable cap
 
        return Response([
            {
                'id': m.id,
                'message': m.message,
                'message_type': m.message_type,
                'created_at': m.created_at.isoformat(),
                'expires_at': m.expires_at.isoformat(),
                'is_active': m.is_active,
                'is_expired': m.expires_at <= now,
                'posted_by': m.created_by.email if m.created_by else 'Unknown',
            }
            for m in messages
        ])
 
 
class AdminDeactivateMessageView(APIView):
    """
    Admin takes a message down early, before its natural 30-day expiry.
    Soft-deactivation (is_active=False), not a hard delete — keeps the
    record around for the admin's own history view.
    URL: PATCH /api/admin/messages/<id>/deactivate/
    """
    permission_classes = [IsAuthenticated]
 
    def patch(self, request, message_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        try:
            msg = SystemMessage.objects.get(id=message_id)
        except SystemMessage.DoesNotExist:
            return Response({'detail': 'Message not found.'}, status=status.HTTP_404_NOT_FOUND)
 
        msg.is_active = False
        msg.save(update_fields=['is_active'])
 
        return Response({'detail': 'Message removed.'})
 
 
class ActiveMessagesView(APIView):
    """
    Any authenticated user (not just admins) — the messages actually shown
    in the sidebar. Only active, unexpired messages, newest first.
    URL: GET /api/messages/active/
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request):
        now = timezone.now()
        messages = SystemMessage.objects.filter(is_active=True, expires_at__gt=now)[:10]
 
        return Response([
            {
                'id': m.id,
                'message': m.message,
                'message_type': m.message_type,
                'created_at': m.created_at.isoformat(),
            }
            for m in messages
        ])

def generate_temp_password(length: int = 12) -> str:
    """Generates a random, readable temp password for trial account handoff."""
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))
 
 
class AdminCreateTrialAccountView(APIView):
    """
    Creates a temporary trial account: a dedicated Organization (is_trial=True,
    subscription_plan='trial') plus one CustomUser under it. The user is
    capped to exactly 1 analysis via TIER_UPLOAD_LIMITS in ai_service/views.py
    (trial org still needs to exist for quota enforcement to apply — a user
    with organization=None bypasses quota checks entirely, which is how the
    system admin account works, so trial users need their own org).
 
    The whole org+user gets hard-deleted once trial_expires_at passes,
    regardless of usage — see cleanup_expired_trials management command.
 
    Returns the generated password ONCE, in this response — it is never
    stored in plaintext or retrievable again afterward. Share it with the
    prospective client directly (call, email, etc.) — this endpoint does not
    email it automatically.
    URL: POST /api/admin/create-trial/
    """
    permission_classes = [IsAuthenticated]
 
    TRIAL_DURATION_HOURS = 48
 
    def post(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        email = request.data.get('email', '').strip()
        first_name = request.data.get('first_name', '').strip()
        last_name = request.data.get('last_name', '').strip()
 
        if not all([email, first_name, last_name]):
            return Response(
                {'detail': 'first_name, last_name, and email are all required.'},
                status=status.HTTP_400_BAD_REQUEST
            )
 
        if CustomUser.objects.filter(email=email).exists():
            return Response({'detail': 'A user with this email already exists.'}, status=status.HTTP_400_BAD_REQUEST)
 
        org_name = f"Trial - {first_name} {last_name}"
        expires_at = timezone.now() + timedelta(hours=self.TRIAL_DURATION_HOURS)
 
        org = Organization.objects.create(
            name=org_name,
            primary_contact_name=first_name,
            primary_contact_surname=last_name,
            primary_email=email,
            subscription_plan='trial',
            team_size_limit=1,
            account_status='active',
            is_paid=False,
            is_trial=True,
            trial_expires_at=expires_at,
        )
 
        temp_password = generate_temp_password()
 
        user = CustomUser.objects.create_user(
            email=email,
            first_name=first_name,
            last_name=last_name,
            password=temp_password,
            organization=org,
            role='member',
        )
 
        log_activity(
            user=request.user,
            action='CREATE_USER',
            target=user.email,
            details=f"Created trial account, expires {expires_at.isoformat()}",
            ip_address=request.META.get('REMOTE_ADDR')
        )
 
        return Response({
            'detail': f'Trial account created for {email}.',
            'email': user.email,
            'temp_password': temp_password,
            'expires_at': expires_at.isoformat(),
            'organization_id': org.id,
        }, status=status.HTTP_201_CREATED)

class UserReportIssueView(APIView):
    """
    Any authenticated user submits an issue report.
    URL: POST /api/issues/report/
    """
    permission_classes = [IsAuthenticated]
 
    def post(self, request):
        category = request.data.get('category', 'bug')
        description = request.data.get('description', '').strip()
 
        if not description:
            return Response({'detail': 'Description is required.'}, status=status.HTTP_400_BAD_REQUEST)
 
        if category not in dict(IssueReport.CATEGORY_CHOICES):
            return Response({'detail': 'Invalid category.'}, status=status.HTTP_400_BAD_REQUEST)
 
        IssueReport.objects.create(
            reported_by=request.user,
            organization=getattr(request.user, 'organization', None),
            category=category,
            description=description,
        )
 
        return Response({'detail': "Thanks — your report has been received."}, status=status.HTTP_201_CREATED)
 
 
class AdminIssueListView(APIView):
    """
    Admin's issue review page — all reports, newest first.
    URL: GET /api/admin/issues/
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        issues = IssueReport.objects.all()[:200]
 
        return Response([
            {
                'id': i.id,
                'category': i.category,
                'category_display': i.get_category_display(),
                'description': i.description,
                'status': i.status,
                'reported_by': i.reported_by.email if i.reported_by else 'Unknown (deleted)',
                'organization': i.organization.name if i.organization else 'Unknown',
                'created_at': i.created_at.isoformat(),
                'resolved_at': i.resolved_at.isoformat() if i.resolved_at else None,
            }
            for i in issues
        ])
 
 
class AdminToggleIssueStatusView(APIView):
    """
    Toggles an issue between open and resolved.
    URL: PATCH /api/admin/issues/<id>/toggle-status/
    """
    permission_classes = [IsAuthenticated]
 
    def patch(self, request, issue_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
 
        try:
            issue = IssueReport.objects.get(id=issue_id)
        except IssueReport.DoesNotExist:
            return Response({'detail': 'Issue not found.'}, status=status.HTTP_404_NOT_FOUND)
 
        if issue.status == 'open':
            issue.status = 'resolved'
            issue.resolved_at = timezone.now()
        else:
            issue.status = 'open'
            issue.resolved_at = None
 
        issue.save(update_fields=['status', 'resolved_at'])
 
        return Response({
            'id': issue.id,
            'status': issue.status,
            'resolved_at': issue.resolved_at.isoformat() if issue.resolved_at else None,
        })


# ==========================================
# CUSTOM EXCEL TEMPLATE (per-organization) — see
# ai_service/custom_template_generator.py for the generation side and
# ai_service/models.py's ClientExcelTemplate docstring for the storage
# model. These three endpoints are the only way a template gets onto an
# org (no self-serve upload — admin-provisioned, same as account creation
# itself).
# ==========================================

# Hard ceiling on an uploaded template's size. Not a fabricated business
# rule — a plain infrastructure safety limit so a mistaken upload (or an
# abusive one) can't fill the VPS's private_media volume. 20MB comfortably
# covers a real master valuation workbook with embedded branding/macros.
MAX_TEMPLATE_UPLOAD_BYTES = 20 * 1024 * 1024

ALLOWED_TEMPLATE_EXTENSIONS = ('.xlsx', '.xlsm')


class AdminUploadExcelTemplateView(APIView):
    """
    Uploads (or replaces) an organization's custom Excel template.
    URL: POST /api/admin/organizations/<org_id>/excel-template/upload/
    Body: multipart/form-data, field 'file'.

    The uploaded bytes are stored EXACTLY as received -- never re-saved
    through openpyxl -- so the client's native formulas, macros, and
    branding are preserved byte-for-byte. openpyxl is only used here to
    read the workbook's defined names for validation; that in-memory
    workbook is discarded after validate_template() runs.
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        upload = request.FILES.get('file')
        if not upload:
            return Response({'detail': "No file provided (expected form field 'file')."}, status=status.HTTP_400_BAD_REQUEST)

        original_filename = upload.name
        if not original_filename.lower().endswith(ALLOWED_TEMPLATE_EXTENSIONS):
            return Response(
                {'detail': f"Unsupported file type. Expected one of: {', '.join(ALLOWED_TEMPLATE_EXTENSIONS)}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        if upload.size > MAX_TEMPLATE_UPLOAD_BYTES:
            return Response(
                {'detail': f"File is too large ({upload.size / (1024*1024):.1f}MB). Max allowed is "
                           f"{MAX_TEMPLATE_UPLOAD_BYTES // (1024*1024)}MB."},
                status=status.HTTP_400_BAD_REQUEST
            )

        file_bytes = upload.read()

        try:
            workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        except Exception as e:
            return Response(
                {'detail': f"Couldn't open this file as an Excel workbook: {e}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        try:
            report = validate_template(workbook)
        except TemplateValidationError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)
        finally:
            workbook.close()

        # Replace any existing template for this org. Deleting the old
        # physical file FIRST matters: client_template_upload_path() keys
        # the filename only by organization id, not by the client's own
        # filename, so without this the storage backend would silently
        # append a "_1" suffix to the new file rather than overwrite the
        # old one -- leaving an orphaned stale template on disk.
        try:
            existing = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            existing = None

        if existing is not None and existing.file:
            existing.file.delete(save=False)
            template = existing
        else:
            template = ClientExcelTemplate(organization=org)

        template.file.save(original_filename, ContentFile(file_bytes), save=False)
        template.original_filename = original_filename
        template.detected_named_ranges = report['found_fields']
        template.uploaded_by = request.user
        template.is_active = True
        template.save()

        return Response({
            'detail': f"Custom template uploaded for {org.name}.",
            'organization_id': org.id,
            'original_filename': template.original_filename,
            'uploaded_at': template.uploaded_at.isoformat(),
            'found_fields': report['found_fields'],
            'missing_scalar_fields': report['missing_scalar_fields'],
            'missing_row_fields': report['missing_row_fields'],
            'row_anchor': report['row_anchor'],
            'unresolvable_defined_names': report['unresolvable_defined_names'],
        }, status=status.HTTP_201_CREATED)


class AdminGetExcelTemplateView(APIView):
    """
    Reports whether an organization has an active custom template, and
    what our engine can/can't do with it, without re-opening the workbook
    (detected_named_ranges was captured once, at upload time).
    URL: GET /api/admin/organizations/<org_id>/excel-template/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        try:
            template = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            return Response({'has_template': False})

        found = set(template.detected_named_ranges or [])
        missing_scalar = sorted(set(SCALAR_FIELD_CATALOG) - found)
        missing_row = sorted(set(ROW_FIELD_CATALOG) - found)

        return Response({
            'has_template': True,
            'is_active': template.is_active,
            'original_filename': template.original_filename,
            'uploaded_at': template.uploaded_at.isoformat(),
            'uploaded_by': template.uploaded_by.email if template.uploaded_by else None,
            'detected_named_ranges': sorted(found),
            'missing_scalar_fields': missing_scalar,
            'missing_row_fields': missing_row,
        })


class AdminDeleteExcelTemplateView(APIView):
    """
    Removes an organization's custom template (file + record). The org
    falls back to the standard generator on its very next export --
    UnderwritePropertyDownloadView looks this up fresh on every request,
    no caching/flag to clear elsewhere.
    URL: DELETE /api/admin/organizations/<org_id>/excel-template/delete/
    """
    permission_classes = [IsAuthenticated]

    def delete(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        try:
            template = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            return Response({'detail': 'This organization has no custom template.'}, status=status.HTTP_404_NOT_FOUND)

        if template.file:
            template.file.delete(save=False)
        template.delete()

        return Response({'detail': f"Custom template removed for {org.name}."}, status=status.HTTP_204_NO_CONTENT)


class MyOrganizationExcelTemplateView(APIView):
    """
    Self-serve counterpart to the three Admin*ExcelTemplate views above,
    scoped to the CALLING USER's own organization — no org_id in the URL
    at all, so there is no way for one org's admin to ever reach another
    org's template by guessing/passing a different id. This is the first
    endpoint in the codebase that grants an org's OWN admin (role='admin'
    on CustomUser) write access to something, as distinct from
    request.user.is_admin, which is the hardcoded platform-superuser flag
    (see CustomUser.save()) used by every Admin* view above — those are
    unchanged and remain for platform support/ops use.

    GET is intentionally open to ANY authenticated member of the org, not
    just its admin: a regular team member needs to know whether a custom
    template exists (and whether it's active) to make sense of the
    "System Standard vs. Custom Template" choice in the upload flow. Only
    POST (upload/replace) and DELETE require the org's own admin role —
    installing or removing the org's proprietary template is treated the
    same way as any other organization-level setting.

    Deliberately reuses validate_template()/ClientExcelTemplate exactly as
    the platform-admin views do — no new matching/validation logic, per
    the decision to keep the named-range engine exactly as-is and only
    change who can reach it and from where.
    """
    permission_classes = [IsAuthenticated]

    def _forbidden_unless_org_admin(self, request):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        if request.user.role != 'admin':
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
        return None

    def get(self, request):
        org = request.user.organization
        if org is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)

        try:
            template = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            return Response({'has_template': False})

        found = set(template.detected_named_ranges or [])
        missing_scalar = sorted(set(SCALAR_FIELD_CATALOG) - found)
        missing_row = sorted(set(ROW_FIELD_CATALOG) - found)

        return Response({
            'has_template': True,
            'is_active': template.is_active,
            'original_filename': template.original_filename,
            'uploaded_at': template.uploaded_at.isoformat(),
            'uploaded_by': template.uploaded_by.email if template.uploaded_by else None,
            'detected_named_ranges': sorted(found),
            'missing_scalar_fields': missing_scalar,
            'missing_row_fields': missing_row,
        })

    def post(self, request):
        denied = self._forbidden_unless_org_admin(request)
        if denied:
            return denied
        org = request.user.organization

        upload = request.FILES.get('file')
        if not upload:
            return Response({'detail': "No file provided (expected form field 'file')."}, status=status.HTTP_400_BAD_REQUEST)

        original_filename = upload.name
        if not original_filename.lower().endswith(ALLOWED_TEMPLATE_EXTENSIONS):
            return Response(
                {'detail': f"Unsupported file type. Expected one of: {', '.join(ALLOWED_TEMPLATE_EXTENSIONS)}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        if upload.size > MAX_TEMPLATE_UPLOAD_BYTES:
            return Response(
                {'detail': f"File is too large ({upload.size / (1024*1024):.1f}MB). Max allowed is "
                           f"{MAX_TEMPLATE_UPLOAD_BYTES // (1024*1024)}MB."},
                status=status.HTTP_400_BAD_REQUEST
            )

        file_bytes = upload.read()

        try:
            workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        except Exception as e:
            return Response(
                {'detail': f"Couldn't open this file as an Excel workbook: {e}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        try:
            report = validate_template(workbook)
        except TemplateValidationError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)
        finally:
            workbook.close()

        try:
            existing = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            existing = None

        if existing is not None and existing.file:
            existing.file.delete(save=False)
            template = existing
        else:
            template = ClientExcelTemplate(organization=org)

        template.file.save(original_filename, ContentFile(file_bytes), save=False)
        template.original_filename = original_filename
        template.detected_named_ranges = report['found_fields']
        template.uploaded_by = request.user
        template.is_active = True
        template.save()

        return Response({
            'detail': "Custom template uploaded.",
            'original_filename': template.original_filename,
            'uploaded_at': template.uploaded_at.isoformat(),
            'found_fields': report['found_fields'],
            'missing_scalar_fields': report['missing_scalar_fields'],
            'missing_row_fields': report['missing_row_fields'],
            'row_anchor': report['row_anchor'],
            'unresolvable_defined_names': report['unresolvable_defined_names'],
        }, status=status.HTTP_201_CREATED)

    def delete(self, request):
        denied = self._forbidden_unless_org_admin(request)
        if denied:
            return denied
        org = request.user.organization

        try:
            template = ClientExcelTemplate.objects.get(organization=org)
        except ClientExcelTemplate.DoesNotExist:
            return Response({'detail': 'Your organization has no custom template.'}, status=status.HTTP_404_NOT_FOUND)

        if template.file:
            template.file.delete(save=False)
        template.delete()

        return Response({'detail': "Custom template removed."}, status=status.HTTP_204_NO_CONTENT)


# ---------------------------------------------------------------------------
# Feature 3 (Enterprise Pipeline & Workflow Dashboard): Team Workspace
# management. Two parallel surfaces, same pattern used for Excel templates
# above — a self-serve one scoped to the calling org-admin's own
# organization (My*), and a platform-admin mirror scoped by org_id in the
# URL (Admin*) for when YOU need to set this up on a client's behalf. Both
# ultimately touch the same TeamWorkspace/CustomUser rows.
# ---------------------------------------------------------------------------

def _serialize_workspace(workspace):
    return {
        'id': workspace.id,
        'name': workspace.name,
        'member_count': workspace.members.count(),
        'project_count': workspace.property_projects.filter(is_deleted=False).count(),
        'created_at': workspace.created_at.isoformat(),
    }


class MyOrganizationWorkspacesView(APIView):
    """
    Self-serve Team Workspace list/create, scoped to the calling user's own
    organization — mirrors MyOrganizationExcelTemplateView's pattern
    exactly: GET open to any authenticated member of the org (everyone
    benefits from seeing what workspaces exist), POST (create) gated to the
    org's own admin (role='admin').
    """
    permission_classes = [IsAuthenticated]

    def _forbidden_unless_org_admin(self, request):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        if request.user.role != 'admin':
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
        return None

    def get(self, request):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        workspaces = TeamWorkspace.objects.filter(organization_id=request.user.organization_id)
        return Response([_serialize_workspace(w) for w in workspaces])

    def post(self, request):
        denied = self._forbidden_unless_org_admin(request)
        if denied:
            return denied

        name = (request.data.get('name') or '').strip()
        if not name:
            return Response({'detail': 'A workspace name is required.'}, status=status.HTTP_400_BAD_REQUEST)

        if TeamWorkspace.objects.filter(organization_id=request.user.organization_id, name=name).exists():
            return Response(
                {'detail': f"A workspace named '{name}' already exists."},
                status=status.HTTP_400_BAD_REQUEST
            )

        workspace = TeamWorkspace.objects.create(
            organization=request.user.organization,
            name=name,
            created_by=request.user,
        )
        return Response(_serialize_workspace(workspace), status=status.HTTP_201_CREATED)


class MyOrganizationWorkspaceDetailView(APIView):
    """Rename/delete ONE workspace within the calling org-admin's own org."""
    permission_classes = [IsAuthenticated]

    def _get_workspace_or_404(self, request, workspace_id):
        try:
            return TeamWorkspace.objects.get(pk=workspace_id, organization_id=request.user.organization_id), None
        except TeamWorkspace.DoesNotExist:
            return None, Response({'detail': 'Workspace not found.'}, status=status.HTTP_404_NOT_FOUND)

    def _forbidden_unless_org_admin(self, request):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        if request.user.role != 'admin':
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
        return None

    def patch(self, request, workspace_id):
        denied = self._forbidden_unless_org_admin(request)
        if denied:
            return denied

        workspace, error = self._get_workspace_or_404(request, workspace_id)
        if error:
            return error

        name = (request.data.get('name') or '').strip()
        if not name:
            return Response({'detail': 'A workspace name is required.'}, status=status.HTTP_400_BAD_REQUEST)

        if TeamWorkspace.objects.filter(
            organization_id=request.user.organization_id, name=name
        ).exclude(pk=workspace.pk).exists():
            return Response(
                {'detail': f"A workspace named '{name}' already exists."},
                status=status.HTTP_400_BAD_REQUEST
            )

        workspace.name = name
        workspace.save(update_fields=['name'])
        return Response(_serialize_workspace(workspace))

    def delete(self, request, workspace_id):
        denied = self._forbidden_unless_org_admin(request)
        if denied:
            return denied

        workspace, error = self._get_workspace_or_404(request, workspace_id)
        if error:
            return error

        # Both CustomUser.workspace and PropertyProject.workspace use
        # SET_NULL (see TeamWorkspace's own docstring) — deleting a
        # workspace never cascades into deleting members OR deals, it just
        # un-assigns them (members fall back to the own-created-projects-
        # only visibility rule; projects become visible only to the org
        # admin and whoever created them).
        workspace.delete()
        return Response({'detail': 'Workspace deleted.'})


class MyOrganizationMembersView(APIView):
    """
    Self-serve roster of the calling org-admin's own team, with each
    member's current workspace — the list the assignment UI picks from.
    Org-admin only (unlike MyOrganizationExcelTemplateView's GET, this
    reveals every teammate's info, not just a yes/no status — a regular
    member doesn't need this).
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        if request.user.role != 'admin':
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        members = CustomUser.objects.filter(
            organization_id=request.user.organization_id
        ).order_by('first_name', 'last_name')

        return Response([
            {
                'email': m.email,
                'first_name': m.first_name,
                'last_name': m.last_name,
                'role': m.role,
                'is_active': m.is_active,
                'workspace_id': m.workspace_id,
                'workspace_name': m.workspace.name if m.workspace_id else None,
                'workspace_role': m.workspace_role,
            }
            for m in members
        ])


class MyOrganizationMemberWorkspaceView(APIView):
    """
    Org-admin assigns (or clears, via workspace_id=null) ONE of their own
    team member's workspace. Enforces the same-organization constraint at
    this layer via AssignUserWorkspaceSerializer — see TeamWorkspace's own
    docstring for why the database itself can't express it.
    """
    permission_classes = [IsAuthenticated]

    def patch(self, request, email):
        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)
        if request.user.role != 'admin':
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        decoded_email = urllib.parse.unquote(email)
        try:
            member = CustomUser.objects.get(email=decoded_email, organization_id=request.user.organization_id)
        except CustomUser.DoesNotExist:
            return Response({'detail': 'Team member not found in your organization.'}, status=status.HTTP_404_NOT_FOUND)

        serializer = AssignUserWorkspaceSerializer(
            data=request.data, context={'organization': request.user.organization}
        )
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        # Feature 4: apply each field only if it was actually present in this
        # request. Both fields are optional on the serializer so a caller can
        # PATCH workspace_role alone (e.g. promoting someone to Head without
        # re-sending their existing workspace_id) - falling back to .get()
        # with a None default here would silently unassign the workspace on
        # a workspace_role-only request.
        update_fields = []
        if 'workspace_id' in serializer.validated_data:
            member.workspace_id = serializer.validated_data['workspace_id']
            update_fields.append('workspace')
        if 'workspace_role' in serializer.validated_data:
            member.workspace_role = serializer.validated_data['workspace_role']
            update_fields.append('workspace_role')

        if update_fields:
            member.save(update_fields=update_fields)

        return Response({
            'email': member.email,
            'workspace_id': member.workspace_id,
            'workspace_name': member.workspace.name if member.workspace_id else None,
            'workspace_role': member.workspace_role,
        })


class AdminOrganizationWorkspacesView(APIView):
    """Platform-admin mirror of MyOrganizationWorkspacesView, scoped by org_id in the URL."""
    permission_classes = [IsAuthenticated]

    def get(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        workspaces = TeamWorkspace.objects.filter(organization=org)
        return Response([_serialize_workspace(w) for w in workspaces])

    def post(self, request, org_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)
        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        name = (request.data.get('name') or '').strip()
        if not name:
            return Response({'detail': 'A workspace name is required.'}, status=status.HTTP_400_BAD_REQUEST)
        if TeamWorkspace.objects.filter(organization=org, name=name).exists():
            return Response(
                {'detail': f"A workspace named '{name}' already exists."},
                status=status.HTTP_400_BAD_REQUEST
            )

        workspace = TeamWorkspace.objects.create(organization=org, name=name, created_by=request.user)
        return Response(_serialize_workspace(workspace), status=status.HTTP_201_CREATED)


class AdminOrganizationWorkspaceDetailView(APIView):
    """Platform-admin mirror of MyOrganizationWorkspaceDetailView, scoped by org_id in the URL."""
    permission_classes = [IsAuthenticated]

    def _get_workspace_or_404(self, org_id, workspace_id):
        try:
            return TeamWorkspace.objects.get(pk=workspace_id, organization_id=org_id), None
        except TeamWorkspace.DoesNotExist:
            return None, Response({'detail': 'Workspace not found.'}, status=status.HTTP_404_NOT_FOUND)

    def patch(self, request, org_id, workspace_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        workspace, error = self._get_workspace_or_404(org_id, workspace_id)
        if error:
            return error

        name = (request.data.get('name') or '').strip()
        if not name:
            return Response({'detail': 'A workspace name is required.'}, status=status.HTTP_400_BAD_REQUEST)
        if TeamWorkspace.objects.filter(organization_id=org_id, name=name).exclude(pk=workspace.pk).exists():
            return Response(
                {'detail': f"A workspace named '{name}' already exists."},
                status=status.HTTP_400_BAD_REQUEST
            )

        workspace.name = name
        workspace.save(update_fields=['name'])
        return Response(_serialize_workspace(workspace))

    def delete(self, request, org_id, workspace_id):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        workspace, error = self._get_workspace_or_404(org_id, workspace_id)
        if error:
            return error

        workspace.delete()
        return Response({'detail': 'Workspace deleted.'})


class AdminOrganizationUserWorkspaceView(APIView):
    """Platform-admin mirror of MyOrganizationMemberWorkspaceView, scoped by org_id in the URL."""
    permission_classes = [IsAuthenticated]

    def patch(self, request, org_id, email):
        if not request.user.is_admin:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        try:
            org = Organization.objects.get(id=org_id)
        except Organization.DoesNotExist:
            return Response({'detail': 'Organization not found'}, status=status.HTTP_404_NOT_FOUND)

        decoded_email = urllib.parse.unquote(email)
        try:
            member = CustomUser.objects.get(email=decoded_email, organization=org)
        except CustomUser.DoesNotExist:
            return Response({'detail': 'User not found in this organization.'}, status=status.HTTP_404_NOT_FOUND)

        serializer = AssignUserWorkspaceSerializer(data=request.data, context={'organization': org})
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        # See MyOrganizationMemberWorkspaceView.patch - same reasoning for
        # only applying a field that was actually present in the request.
        update_fields = []
        if 'workspace_id' in serializer.validated_data:
            member.workspace_id = serializer.validated_data['workspace_id']
            update_fields.append('workspace')
        if 'workspace_role' in serializer.validated_data:
            member.workspace_role = serializer.validated_data['workspace_role']
            update_fields.append('workspace_role')

        if update_fields:
            member.save(update_fields=update_fields)

        return Response({
            'email': member.email,
            'workspace_id': member.workspace_id,
            'workspace_name': member.workspace.name if member.workspace_id else None,
            'workspace_role': member.workspace_role,
        })


class WorkspaceUsageView(APIView):
    """
    Feature 4 ("Head of Underwriting" role hierarchy): workspace-scoped
    usage metrics — the Head Seat's "clear view of their specific team's
    activity" the stakeholder brief asked for.

    Scoped by PropertyProject.workspace_id, NOT by which members currently
    have this workspace set on their account — the same choice
    ai_service/permissions.py's visible_projects_for makes for a workspace
    Head's deal visibility, and for the same reason: a deal's workspace is
    a property of the PROJECT. That keeps this view honest in the two cases
    where "who's in the workspace today" and "whose work happened in the
    workspace" diverge: a member who created deals here and has since been
    moved to a different workspace (or had their workspace cleared) still
    shows up against the deals they created while they were here, and a
    project that's since been reassigned to a different workspace moves
    with it rather than staying attributed to this one.

    URL: GET /api/organization/workspaces/<int:workspace_id>/usage/
    Gated to: the org's own admin (role='admin'), OR a member of THIS
    workspace with workspace_role='head'. Never any other workspace's head
    — matching the "no cross-workspace access, ever" decision locked in
    for this feature.
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, workspace_id):
        from django.db.models import Count

        if request.user.organization_id is None:
            return Response({'detail': 'This account has no organization.'}, status=status.HTTP_400_BAD_REQUEST)

        try:
            workspace = TeamWorkspace.objects.get(pk=workspace_id, organization_id=request.user.organization_id)
        except TeamWorkspace.DoesNotExist:
            return Response({'detail': 'Workspace not found.'}, status=status.HTTP_404_NOT_FOUND)

        is_this_workspaces_head = (
            request.user.workspace_id == workspace.id and request.user.workspace_role == 'head'
        )
        if request.user.role != 'admin' and not is_this_workspaces_head:
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        projects = PropertyProject.objects.filter(workspace=workspace, is_deleted=False)

        deals_by_status = {choice_value: 0 for choice_value, _ in PropertyProject.STATUS_CHOICES}
        for row in projects.values('deal_status').annotate(count=Count('id')):
            deals_by_status[row['deal_status']] = row['count']

        analyses = AnalysisReport.objects.filter(project__workspace=workspace, is_deleted=False)

        now = timezone.now()
        thirty_days_ago = now - timedelta(days=30)

        members = CustomUser.objects.filter(workspace=workspace).order_by('first_name', 'last_name')

        per_member = [
            {
                'email': member.email,
                'first_name': member.first_name,
                'last_name': member.last_name,
                'workspace_role': member.workspace_role,
                'deals_created': projects.filter(created_by_id=member.pk).count(),
                'analyses_run': analyses.filter(uploaded_by_id=member.pk).count(),
            }
            for member in members
        ]

        return Response({
            'workspace_id': workspace.id,
            'workspace_name': workspace.name,
            'member_count': members.count(),
            'deals_total': projects.count(),
            'deals_by_status': deals_by_status,
            'analyses_total': analyses.count(),
            'analyses_last_30_days': analyses.filter(created_at__gte=thirty_days_ago).count(),
            'members': per_member,
        })