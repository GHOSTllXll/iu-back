from django.shortcuts import render

# Create your views here.
from django.core.mail import EmailMessage
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status


class ContactView(APIView):

    def post(self, request):

        company = request.data.get("company")
        company_size = request.data.get("companySize")
        contact = request.data.get("contact")
        email = request.data.get("email")
        phone = request.data.get("phone")
        message = request.data.get("message")

        body = f"""
New Website Enquiry

Company:
{company}

Company Size:
{company_size}

Contact:
{contact}

Email:
{email}

Phone:
{phone}

Message:
{message}
"""

        EmailMessage(
            subject=f"Website Enquiry - {company}",
            body=body,
            from_email="donovan@inboundunderwriting.com",
            to=["donovan@inboundunderwriting.com"],
            reply_to=[email] if email else [],
        ).send()

        return Response(
            {"success": True},
            status=status.HTTP_200_OK,
        )