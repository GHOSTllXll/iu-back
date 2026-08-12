# backend/users/urls.py
from django.urls import path, include
from .views import (LoginView, LogoutView, CurrentUserView, AdminCreateUserView, AdminOrganizationsView,
                    AdminOrganizationUsersView, AdminToggleUserStatusView, AdminToggleOrganizationStatusView, AdminDeleteOrganizationView,
                    AdminDeleteUserView, AdminActiveUsersCountView, AdminActivityLogView, AdminUpdateOrganizationPlanView,
                    PasswordResetRequestView, PasswordResetConfirmView,
                    AdminCreateMessageView, AdminMessageListView, AdminDeactivateMessageView, ActiveMessagesView,
                    AdminCreateTrialAccountView,
                    UserReportIssueView, AdminIssueListView, AdminToggleIssueStatusView)
from ai_service.urls import urlpatterns as ai_urls

urlpatterns = [
    path('auth/login/', LoginView.as_view(), name='login'),
    path('auth/logout/', LogoutView.as_view(), name='logout'),

    path('auth/user/', CurrentUserView.as_view(), name='current_user'),

    path('auth/password-reset/request/', PasswordResetRequestView.as_view(), name='password_reset_request'),
    path('auth/password-reset/confirm/', PasswordResetConfirmView.as_view(), name='password_reset_confirm'),

    path('admin/create-user/', AdminCreateUserView.as_view(), name='admin_create_user'),
    path('admin/create-trial/', AdminCreateTrialAccountView.as_view(), name='admin_create_trial'),
    path('admin/organizations/', AdminOrganizationsView.as_view(), name='admin_organizations'),

    path('admin/organizations/<int:org_id>/users/', AdminOrganizationUsersView.as_view(), name='admin_org_users'),
    path('admin/users/<str:email>/toggle-status/', AdminToggleUserStatusView.as_view(), name='admin_toggle_user'),

    path('admin/organizations/<int:org_id>/toggle-status/', AdminToggleOrganizationStatusView.as_view(), name='admin_toggle_org_status'),
    path('admin/organizations/<int:org_id>/delete/', AdminDeleteOrganizationView.as_view(), name='admin_delete_org'),

    path('admin/users/<str:email>/delete/', AdminDeleteUserView.as_view(), name='admin_delete_user'),

    path('admin/active-users-count/', AdminActiveUsersCountView.as_view(), name='admin_active_users_count'),

    path('admin/activity-logs/', AdminActivityLogView.as_view(), name='admin_activity_logs'),

    path('admin/organizations/<int:org_id>/update-plan/', AdminUpdateOrganizationPlanView.as_view(), name='admin_update_org_plan'),

    path('admin/messages/', AdminMessageListView.as_view(), name='admin_message_list'),
    path('admin/messages/create/', AdminCreateMessageView.as_view(), name='admin_message_create'),
    path('admin/messages/<int:message_id>/deactivate/', AdminDeactivateMessageView.as_view(), name='admin_message_deactivate'),
    path('messages/active/', ActiveMessagesView.as_view(), name='active_messages'),

    # Issue reports
    path('issues/report/', UserReportIssueView.as_view(), name='user_report_issue'),
    path('admin/issues/', AdminIssueListView.as_view(), name='admin_issue_list'),
    path('admin/issues/<int:issue_id>/toggle-status/', AdminToggleIssueStatusView.as_view(), name='admin_toggle_issue_status'),

    path('ai/', include(ai_urls)),

]