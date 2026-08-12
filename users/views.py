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
from .models import CustomUser, Organization, ActivityLog, SystemMessage, IssueReport
from .serializers import LoginSerializer, UserSerializer, AdminCreateUserSerializer, OrganizationSerializer, ActivityLogSerializer

from django.contrib.auth.tokens import PasswordResetTokenGenerator
from django.utils.encoding import force_bytes, force_str
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.core.mail import send_mail
from django.conf import settings

from rest_framework.throttling import AnonRateThrottle, SimpleRateThrottle

import secrets
import string

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
 
    TRIAL_DURATION_HOURS = 24
 
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