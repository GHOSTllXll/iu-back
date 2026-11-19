"""
URL configuration for config project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include

# Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): imported
# directly from ai_service.views rather than routed through
# ai_service/urls.py's own urlpatterns (which only ever get included under
# users.urls' 'ai/' prefix, i.e. /api/ai/...). This endpoint is a
# server-to-server webhook Postmark posts to - never this app's own
# frontend - so it deliberately sits at its own /api/webhooks/... path
# instead of inheriting the /api/ai/ namespace that exists for this
# project's authenticated browser-facing API.
from ai_service.views import InboundEmailWebhookView

urlpatterns = [
    path('admin/', admin.site.urls),
    path("api/contact/", include("contact.urls")),
    path("api/", include("users.urls")),
    path("api/webhooks/inbound-email/", InboundEmailWebhookView.as_view(), name='inbound_email_webhook'),
]
