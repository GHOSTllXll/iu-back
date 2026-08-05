from .models import ActivityLog

def log_activity(user, action, target, details="", ip_address=""):
    """
    Utility function to quickly log an activity.
    """
    ActivityLog.objects.create(
        user=user,
        action=action,
        target=target,
        details=details,
        ip_address=ip_address
    )