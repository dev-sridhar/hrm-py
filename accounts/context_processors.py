from django.utils import timezone
from accounts.models import LeaveRequest, DailyTask, Attendance, UserStatus


def dashboard_notifications(request):
    if not request.user.is_authenticated:
        return {}

    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()

    # Leave notifications (latest 4)
    all_leaves = LeaveRequest.objects.select_related("user", "user__profile").all().order_by("-applied_at")[:4]
    nav_leave_notifications = []
    for l in all_leaves:
        u_name = l.user.get_full_name() or l.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = l.user.profile.avatar.url if (hasattr(l.user, 'profile') and l.user.profile.avatar) else None
        nav_leave_notifications.append({
            "user_name": u_name,
            "user_initials": u_init,
            "avatar_url": avatar_url,
            "leave_type": l.get_leave_type_display(),
            "days_count": l.days_count,
            "start_date": l.start_date.strftime("%d %b"),
            "end_date": l.end_date.strftime("%d %b"),
            "status": l.status,
            "time_str": timezone.localtime(l.applied_at).strftime("%d %b, %I:%M %p") if l.applied_at else "",
            "url": "/applications/leaves/",
        })

    # Task notifications (latest 4)
    all_tasks = DailyTask.objects.select_related("user", "user__profile").all().order_by("-created_at")[:4]
    nav_task_notifications = []
    for t in all_tasks:
        u_name = t.user.get_full_name() or t.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = t.user.profile.avatar.url if (hasattr(t.user, 'profile') and t.user.profile.avatar) else None
        nav_task_notifications.append({
            "title": t.title,
            "user_name": u_name,
            "user_initials": u_init,
            "avatar_url": avatar_url,
            "due_date": t.due_date.strftime("%d %b %Y") if t.due_date else "Today",
            "priority": t.priority,
            "time_str": timezone.localtime(t.created_at).strftime("%d %b") if t.created_at else "",
            "url": "/tasks/",
        })

    # Activity logs (latest 4)
    all_att = Attendance.objects.select_related("user", "user__profile").all().order_by("-updated_at")[:4]
    nav_activity_logs = []
    for a in all_att:
        u_name = a.user.get_full_name() or a.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = a.user.profile.avatar.url if (hasattr(a.user, 'profile') and a.user.profile.avatar) else None
        if a.punch_out:
            act_text = f"Checked out ({a.formatted_duration})"
            time_str = timezone.localtime(a.punch_out).strftime("%d %b, %I:%M %p")
        elif a.punch_in:
            act_text = "Checked in"
            time_str = timezone.localtime(a.punch_in).strftime("%d %b, %I:%M %p")
        else:
            act_text = f"Marked as {a.get_status_display()}"
            time_str = a.date.strftime("%d %b")
        nav_activity_logs.append({
            "user_name": u_name,
            "user_initials": u_init,
            "avatar_url": avatar_url,
            "action": act_text,
            "time_str": time_str,
            "url": "/attendance/",
        })

    total_pending = LeaveRequest.objects.filter(status="pending").count() + DailyTask.objects.filter(status="todo").count()

    return {
        "nav_leave_notifications": nav_leave_notifications,
        "nav_task_notifications": nav_task_notifications,
        "nav_activity_logs": nav_activity_logs,
        "nav_total_badge_count": total_pending,
        "today_attendance": today_attendance,
    }