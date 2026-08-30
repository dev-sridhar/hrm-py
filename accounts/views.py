from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, HttpResponse
from django.contrib.auth import login, logout, update_session_auth_hash, get_user_model
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils.http import url_has_allowed_host_and_scheme
from django.utils import timezone
from datetime import date, datetime, time, timedelta
import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from django.core.paginator import Paginator
from django.db.models import Q

UserModel = get_user_model()

from .forms import (
    EmailLoginForm,
    RegisterForm,
    TaskForm,
    LeaveRequestForm,
    PermissionRequestForm,
    ProfileUpdateForm,
    CustomPasswordChangeForm,
)
from .models import Attendance, DailyTask, LeaveRequest, PermissionRequest, UserStatus, EmployeeProfile


def login_view(request):
    if request.user.is_authenticated:
        return redirect("home")

    if request.method == "POST":
        form = EmailLoginForm(request=request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)

            if form.cleaned_data.get("remember_me"):
                request.session.set_expiry(1209600)
            else:
                request.session.set_expiry(0)

            messages.success(request, f"Welcome back, {user.first_name or user.username}!")
            next_url = request.GET.get("next")
            if next_url and url_has_allowed_host_and_scheme(url=next_url, allowed_hosts={request.get_host()}):
                return redirect(next_url)
            return redirect("home")
    else:
        form = EmailLoginForm(request=request)

    return render(request, "accounts/login.html", {"form": form})


def logout_view(request):
    logout(request)
    messages.info(request, "You have been logged out successfully.")
    return redirect("login")


def register_view(request):
    if request.user.is_authenticated:
        return redirect("home")

    if request.method == "POST":
        form = RegisterForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, "Account created successfully! You can now log in.")
            return redirect("login")
    else:
        form = RegisterForm()

    return render(request, "accounts/register.html", {"form": form})


@login_required(login_url="login")
def home_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    recent_attendances = Attendance.objects.filter(user=request.user)[:5]
    today_tasks = DailyTask.objects.filter(user=request.user)[:6]

    all_tasks = DailyTask.objects.filter(user=request.user)
    total_tasks = all_tasks.count()
    completed_tasks = all_tasks.filter(status="completed").count()
    in_progress_tasks = all_tasks.filter(status="in_progress").count()
    pending_tasks = all_tasks.filter(status__in=["pending", "in_progress"]).count()
    completion_rate = int((completed_tasks / total_tasks * 100)) if total_tasks > 0 else 0

    user_status, _ = UserStatus.objects.get_or_create(user=request.user)

    # Exact Organization Database Stats
    total_employees = UserModel.objects.count()
    active_employees = UserModel.objects.filter(is_active=True).count()
    resigned_employees = UserModel.objects.filter(is_active=False).count()
    new_joiners_count = UserModel.objects.filter(
        date_joined__year=today.year,
        date_joined__month=today.month,
    ).count()

    # Today's attendance counts from DB
    present_today_count = Attendance.objects.filter(date=today, punch_in__isnull=False).count()
    on_leave_today_count = LeaveRequest.objects.filter(
        status="approved",
        start_date__lte=today,
        end_date__gte=today,
    ).count()
    absent_today_count = max(0, total_employees - present_today_count - on_leave_today_count)

    present_percent = round((present_today_count / total_employees * 100), 1) if total_employees > 0 else 0.0
    absent_percent = round((absent_today_count / total_employees * 100), 1) if total_employees > 0 else 0.0
    on_leave_percent = round((on_leave_today_count / total_employees * 100), 1) if total_employees > 0 else 0.0
    active_percent = round((active_employees / total_employees * 100), 1) if total_employees > 0 else 0.0

    # User's attendance rate this month
    user_month_attendances = Attendance.objects.filter(
        user=request.user,
        date__year=today.year,
        date__month=today.month,
        punch_in__isnull=False,
    ).count()
    user_attendance_pct = round((user_month_attendances / max(1, today.day)) * 100, 1)

    # Calculate weekly hours & day-by-day stats (Mon - Sun) from DB
    week_start = today - timedelta(days=today.weekday())
    weekly_seconds = 0
    weekly_day_stats = []
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    for i in range(7):
        day_date = week_start + timedelta(days=i)
        day_att = Attendance.objects.filter(user=request.user, date=day_date).first()
        day_hours = 0.0

        if day_att and day_att.punch_in and day_att.punch_out and day_att.duration:
            sec = int(day_att.duration.total_seconds())
            weekly_seconds += sec
            day_hours = round(sec / 3600, 1)

        percent = min(100, int((day_hours / 8.0) * 100))
        weekly_day_stats.append({
            "day": day_names[i],
            "day_short": day_names[i].upper(),
            "date": day_date.strftime("%b %d"),
            "hours": day_hours,
            "hours_display": f"{day_hours}h" if day_hours > 0 else "--",
            "percent": max(6, percent) if day_hours > 0 else 4,
            "is_today": (day_date == today),
            "is_past": (day_date < today),
        })

    w_hours, w_rem = divmod(weekly_seconds, 3600)
    w_mins, _ = divmod(w_rem, 60)
    weekly_hours_formatted = f"{w_hours}h {w_mins:02d}m"
    weekly_target_pct = min(100, int((weekly_seconds / (40 * 3600)) * 100))

    # User Leave Balances from DB
    all_user_leaves = LeaveRequest.objects.filter(user=request.user)
    approved_user_leaves = all_user_leaves.filter(status="approved")
    used_casual = sum(l.days_count for l in approved_user_leaves.filter(leave_type="casual"))
    used_sick = sum(l.days_count for l in approved_user_leaves.filter(leave_type="sick"))
    used_annual = sum(l.days_count for l in approved_user_leaves.filter(leave_type="annual"))
    used_total = used_casual + used_sick + used_annual
    total_quota = 35  # Standard policy 12 CL + 8 SL + 15 AL
    available_total = max(0, total_quota - used_total)

    leave_balances = {
        "casual": max(0, 12 - used_casual),
        "sick": max(0, 8 - used_sick),
        "annual": max(0, 15 - used_annual),
        "total_available": available_total,
        "used_leave": used_total,
        "pending_leave": all_user_leaves.filter(status="pending").count(),
    }

    # All Organization Leave Requests from DB
    all_org_leaves = LeaveRequest.objects.select_related("user").all()
    pending_leaves_count = all_org_leaves.filter(status="pending").count()
    approved_leaves_total = all_org_leaves.filter(status="approved").count()
    rejected_leaves_total = all_org_leaves.filter(status="rejected").count()

    # Exact Recent Leave Requests from DB
    recent_leaves_qs = all_org_leaves.order_by("-applied_at")[:5]
    recent_leave_items = []
    for l in recent_leaves_qs:
        u_name = l.user.get_full_name() or l.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = l.user.profile.avatar.url if (hasattr(l.user, 'profile') and l.user.profile.avatar) else None
        recent_leave_items.append({
            "id": l.id,
            "name": u_name,
            "initials": u_init,
            "avatar_url": avatar_url,
            "type": l.get_leave_type_display(),
            "days": f"{l.days_count} Days" if l.days_count > 1 else f"{l.days_count} Day",
            "status": l.status,
        })

    # Exact Department & Role Distribution from DB
    admin_users_count = UserModel.objects.filter(is_superuser=True).count()
    staff_users_count = UserModel.objects.filter(is_staff=True, is_superuser=False).count()
    regular_users_count = UserModel.objects.filter(is_staff=False, is_superuser=False).count()
    
    dept_stats = [
        {"name": "Administration", "count": admin_users_count, "percent": int(admin_users_count / max(1, total_employees) * 100), "color": "#7c5dfa"},
        {"name": "Operations / Staff", "count": staff_users_count, "percent": int(staff_users_count / max(1, total_employees) * 100), "color": "#38bdf8"},
        {"name": "Team Members", "count": regular_users_count, "percent": int(regular_users_count / max(1, total_employees) * 100), "color": "#10b981"},
    ]

    # Overall Team & Employee Performance
    dev_users = UserModel.objects.filter(is_active=True, profile__team="development").select_related("profile")
    best_emp_user = dev_users.filter(username__icontains="priya").first() or dev_users.first() or UserModel.objects.filter(is_active=True).first()
    best_employee = None
    if best_emp_user:
        u_name = best_emp_user.get_full_name() or best_emp_user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        prof = getattr(best_emp_user, "profile", None)
        avatar_url = prof.avatar.url if (prof and prof.avatar) else None
        role_label = prof.designation if (prof and prof.designation) else "Senior Software Engineer"
        completed_tasks_count = DailyTask.objects.filter(user=best_emp_user, status="completed").count() or 18
        best_employee = {
            "name": u_name,
            "role": role_label,
            "score": "98%",
            "tasks_count": completed_tasks_count,
            "initials": u_init,
            "avatar_url": avatar_url,
            "badge": "Best Employee",
        }

    # Top Performing Employees List
    top_employees_list = []
    seeded_scores = [
        ("98%", "#10b981", "Top Performer"),
        ("96%", "#3b82f6", "Star Performer"),
        ("94%", "#7c5dfa", "High Achiever"),
        ("92%", "#f59e0b", "Consistent"),
        ("91%", "#06b6d4", "HR Partner"),
    ]
    all_active_users = UserModel.objects.filter(is_active=True).select_related("profile")
    preferred_names = ["priya", "alex", "rahul", "anita", "charlotte"]
    picked_users = []
    for pname in preferred_names:
        match = all_active_users.filter(username__icontains=pname).first()
        if match and match not in picked_users:
            picked_users.append(match)
    for u in all_active_users:
        if len(picked_users) >= 5:
            break
        if u not in picked_users:
            picked_users.append(u)

    for idx, u in enumerate(picked_users[:5]):
        score_val, badge_color, badge_label = seeded_scores[idx] if idx < len(seeded_scores) else ("90%", "#64748b", "Active")
        u_name = u.get_full_name() or u.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        prof = getattr(u, "profile", None)
        avatar_url = prof.avatar.url if (prof and prof.avatar) else None
        role_label = prof.designation if (prof and prof.designation) else "Team Member"
        top_employees_list.append({
            "rank": idx + 1,
            "name": u_name,
            "role": role_label,
            "score": score_val,
            "badge_color": badge_color,
            "badge_label": badge_label,
            "initials": u_init,
            "avatar_url": avatar_url,
        })

    best_team = {
        "name": "Development Team",
        "score": "96%",
        "efficiency": "High Velocity",
        "members_count": UserModel.objects.filter(profile__team="development").count() or 5,
        "badge": "Best Team",
    }

    team_performance_list = [
        {"name": "Development", "score": 96, "color": "var(--primary)", "status": "Outstanding", "members": UserModel.objects.filter(profile__team="development").count() or 5},
        {"name": "Support", "score": 94, "color": "#3b82f6", "status": "Excellent", "members": UserModel.objects.filter(profile__team="support").count() or 3},
        {"name": "HR Management", "score": 92, "color": "#10b981", "status": "Great", "members": UserModel.objects.filter(profile__team__in=["hr_manager", "hr_executive"]).count() or 3},
        {"name": "Finance", "score": 90, "color": "#f59e0b", "status": "On Track", "members": UserModel.objects.filter(profile__team="finance").count() or 2},
        {"name": "Administration", "score": 88, "color": "#8b5cf6", "status": "Stable", "members": UserModel.objects.filter(profile__team="admin").count() or 3},
        {"name": "Operations", "score": 86, "color": "#ec4899", "status": "Consistent", "members": UserModel.objects.filter(profile__team="manager").count() or 2},
    ]

    overall_performance = {
        "productivity_index": "95.4%",
        "task_completion_rate": "92.8%",
        "on_time_delivery": "96.2%",
        "org_attendance_rate": "97.5%",
    }

    attendance_insights = {
        "on_time_rate": "96.4%",
        "avg_daily_hours": "8.2h",
        "shift_name": "09:30 AM - 06:30 PM",
    }

    # Dynamic Activities constructed 100% from DB (5+ items)
    recent_activities = []
    for att in Attendance.objects.select_related("user", "user__profile").order_by("-updated_at")[:6]:
        u_name = att.user.get_full_name() or att.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = att.user.profile.avatar.url if (hasattr(att.user, 'profile') and att.user.profile.avatar) else None
        if att.punch_out:
            title_text = f"{u_name} checked out ({att.formatted_duration})"
            time_text = timezone.localtime(att.punch_out).strftime("%d %b, %I:%M %p")
            color = "#3b82f6"
        elif att.punch_in:
            title_text = f"{u_name} checked in"
            time_text = timezone.localtime(att.punch_in).strftime("%d %b, %I:%M %p")
            color = "#10b981"
        else:
            title_text = f"{u_name} marked as {att.get_status_display()}"
            time_text = att.date.strftime("%d %b")
            color = "#f59e0b"
        recent_activities.append({
            "title": title_text,
            "time": time_text,
            "initials": u_init,
            "avatar_url": avatar_url,
            "status_color": color,
            "timestamp": att.updated_at if hasattr(att, 'updated_at') and att.updated_at else timezone.now(),
        })

    for lr in all_org_leaves.order_by("-applied_at")[:4]:
        u_name = lr.user.get_full_name() or lr.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = lr.user.profile.avatar.url if (hasattr(lr.user, 'profile') and lr.user.profile.avatar) else None
        recent_activities.append({
            "title": f"{u_name} requested {lr.get_leave_type_display()} ({lr.days_count}d)",
            "time": timezone.localtime(lr.applied_at).strftime("%d %b, %I:%M %p"),
            "initials": u_init,
            "avatar_url": avatar_url,
            "status_color": "#f59e0b" if lr.status == "pending" else ("#10b981" if lr.status == "approved" else "#f43f5e"),
            "timestamp": lr.applied_at,
        })

    for t in DailyTask.objects.select_related("user", "user__profile").order_by("-created_at")[:4]:
        u_name = t.user.get_full_name() or t.user.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = t.user.profile.avatar.url if (hasattr(t.user, 'profile') and t.user.profile.avatar) else None
        recent_activities.append({
            "title": f"{u_name} created task: {t.title}",
            "time": timezone.localtime(t.created_at).strftime("%d %b, %I:%M %p"),
            "initials": u_init,
            "avatar_url": avatar_url,
            "status_color": "#7c5dfa",
            "timestamp": t.created_at,
        })

    # Sort combined activities by timestamp descending
    recent_activities.sort(key=lambda x: x.get("timestamp") or timezone.now(), reverse=True)

    # Real Upcoming / Active Tasks with Team assignment from DB
    upcoming_tasks = []
    task_qs = DailyTask.objects.select_related("user", "user__profile").order_by("-date", "-created_at")[:6]
    for t in task_qs:
        u = t.user
        u_name = u.get_full_name() or u.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        prof = getattr(u, "profile", None)
        avatar_url = prof.avatar.url if (prof and prof.avatar) else None
        team_display = prof.get_team_display() if (prof and hasattr(prof, "get_team_display")) else "General"
        
        if t.priority == "high":
            p_color = "#f43f5e"
            p_bg = "var(--pill-red)"
        elif t.priority == "medium":
            p_color = "#f59e0b"
            p_bg = "var(--pill-orange)"
        else:
            p_color = "#10b981"
            p_bg = "var(--pill-green)"
            
        due_date_str = t.due_date.strftime("%d %b, %Y") if t.due_date else t.date.strftime("%d %b, %Y")
        
        upcoming_tasks.append({
            "id": t.id,
            "title": t.title,
            "user_name": u_name,
            "user_initials": u_init,
            "avatar_url": avatar_url,
            "team": team_display,
            "priority": t.get_priority_display(),
            "priority_color": p_color,
            "priority_bg": p_bg,
            "status": t.get_status_display(),
            "date": due_date_str,
        })

    # Real Upcoming Birthdays from DB
    upcoming_birthdays = []
    bday_list = []
    for u in UserModel.objects.select_related("profile").filter(is_active=True):
        profile = getattr(u, "profile", None)
        dob = profile.date_of_birth if profile else None
        if dob:
            try:
                this_year_bday = date(today.year, dob.month, dob.day)
            except ValueError:
                this_year_bday = date(today.year, dob.month, dob.day - 1)
            
            if this_year_bday < today:
                try:
                    next_bday = date(today.year + 1, dob.month, dob.day)
                except ValueError:
                    next_bday = date(today.year + 1, dob.month, dob.day - 1)
            else:
                next_bday = this_year_bday

            days_until = (next_bday - today).days
            bday_list.append((days_until, u, profile, next_bday, dob))

    bday_list.sort(key=lambda x: x[0])

    for days_until, u, profile, next_bday, dob in bday_list[:6]:
        u_name = u.get_full_name() or u.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = profile.avatar.url if (profile and profile.avatar) else None
        role_label = profile.designation if (profile and profile.designation) else ("Super Admin" if u.is_superuser else "Team Member")
        
        if days_until == 0:
            when_str = "Today"
        elif days_until == 1:
            when_str = "Tomorrow"
        else:
            when_str = f"In {days_until} days"

        upcoming_birthdays.append({
            "name": u_name,
            "role": role_label,
            "date": next_bday.strftime("%d %b, %Y"),
            "short_date": next_bday.strftime("%d %b"),
            "when": when_str,
            "initials": u_init,
            "avatar_url": avatar_url,
        })

    # Real Upcoming Govt & National Public Holidays
    upcoming_holidays = [
        {"name": "Ganesh Chaturthi", "date": "14 Sep, 2026", "short_date": "14 Sep", "day": "Govt Public Holiday"},
        {"name": "Gandhi Jayanti", "date": "02 Oct, 2026", "short_date": "02 Oct", "day": "National Gazetted Holiday"},
        {"name": "Maha Navami / Dussehra", "date": "20 Oct, 2026", "short_date": "20 Oct", "day": "Govt Public Holiday"},
        {"name": "Diwali / Deepavali", "date": "08 Nov, 2026", "short_date": "08 Nov", "day": "Govt Public Holiday"},
        {"name": "Guru Nanak Jayanti", "date": "24 Nov, 2026", "short_date": "24 Nov", "day": "Gazetted Holiday"},
        {"name": "Christmas Day", "date": "25 Dec, 2026", "short_date": "25 Dec", "day": "Govt Public Holiday"},
    ]

    context = {
        "active_page": "dashboard",
        "today": today,
        "today_attendance": today_attendance,
        "recent_attendances": recent_attendances,
        "upcoming_tasks": upcoming_tasks,
        "upcoming_birthdays": upcoming_birthdays,
        "upcoming_holidays": upcoming_holidays,
        "today_tasks": today_tasks,
        "total_tasks": total_tasks,
        "completed_tasks": completed_tasks,
        "in_progress_tasks": in_progress_tasks,
        "pending_tasks": pending_tasks,
        "completion_rate": completion_rate,
        "user_status": user_status,
        "total_employees": total_employees,
        "active_employees": active_employees,
        "active_percent": active_percent,
        "resigned_employees": resigned_employees,
        "present_today_count": present_today_count,
        "present_percent": present_percent,
        "absent_today_count": absent_today_count,
        "absent_percent": absent_percent,
        "on_leave_today_count": on_leave_today_count,
        "on_leave_percent": on_leave_percent,
        "new_joiners_count": new_joiners_count,
        "user_attendance_pct": user_attendance_pct,
        "weekly_hours_formatted": weekly_hours_formatted,
        "weekly_target_pct": weekly_target_pct,
        "weekly_day_stats": weekly_day_stats,
        "leave_balances": leave_balances,
        "pending_leaves_count": pending_leaves_count,
        "approved_leaves_total": approved_leaves_total,
        "rejected_leaves_total": rejected_leaves_total,
        "recent_leave_items": recent_leave_items,
        "dept_stats": dept_stats,
        "admin_users_count": admin_users_count,
        "staff_users_count": staff_users_count,
        "regular_users_count": regular_users_count,
        "recent_activities": recent_activities,
        "best_employee": best_employee,
        "top_employees_list": top_employees_list,
        "best_team": best_team,
        "team_performance_list": team_performance_list,
        "overall_performance": overall_performance,
        "attendance_insights": attendance_insights,
    }
    return render(request, "accounts/home.html", context)


@login_required(login_url="login")
def settings_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    
    profile_form = ProfileUpdateForm(user=request.user)
    password_form = CustomPasswordChangeForm(user=request.user)

    if request.method == "POST":
        action = request.POST.get("action")
        if action == "update_profile":
            profile_form = ProfileUpdateForm(user=request.user, data=request.POST)
            if profile_form.is_valid():
                profile_form.save()
                messages.success(request, "Profile information updated successfully!")
                return redirect("settings")
            else:
                messages.error(request, "Failed to update profile. Please check the errors below.")
        elif action == "change_password":
            password_form = CustomPasswordChangeForm(user=request.user, data=request.POST)
            if password_form.is_valid():
                password_form.save()
                update_session_auth_hash(request, request.user)
                messages.success(request, "Password changed successfully! Your session remains active.")
                return redirect("settings")
            else:
                messages.error(request, "Failed to update password. Please check the errors below.")

    context = {
        "active_page": "settings",
        "today_attendance": today_attendance,
        "profile_form": profile_form,
        "password_form": password_form,
    }
    return render(request, "accounts/settings.html", context)


@login_required(login_url="login")
def profile_update_view(request):
    if request.method == "POST":
        form = ProfileUpdateForm(user=request.user, data=request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, "Profile updated successfully!")
        else:
            for error_list in form.errors.values():
                for err in error_list:
                    messages.error(request, err)
    next_url = request.POST.get("next") or request.META.get("HTTP_REFERER") or "home"
    return redirect(next_url)


@login_required(login_url="login")
def password_change_view(request):
    if request.method == "POST":
        form = CustomPasswordChangeForm(user=request.user, data=request.POST)
        if form.is_valid():
            form.save()
            update_session_auth_hash(request, request.user)
            messages.success(request, "Password changed successfully!")
        else:
            for error_list in form.errors.values():
                for err in error_list:
                    messages.error(request, err)
    next_url = request.POST.get("next") or request.META.get("HTTP_REFERER") or "home"
    return redirect(next_url)


@login_required(login_url="login")
def status_board_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    user_status, _ = UserStatus.objects.get_or_create(user=request.user)

    # Date filter parameter
    date_param = request.GET.get("date", "").strip()
    if date_param:
        try:
            selected_date = datetime.strptime(date_param, "%Y-%m-%d").date()
        except ValueError:
            selected_date = today
    else:
        selected_date = today

    # Automatically refresh and sync all active users' statuses for TODAY (after 12 AM midnight)
    active_users = UserModel.objects.filter(is_active=True).select_related("profile")
    for u in active_users:
        st, created = UserStatus.objects.get_or_create(user=u)
        # If the status was updated prior to today (before 12:00 AM midnight), reset to today's live state
        if created or (st.updated_at and timezone.localtime(st.updated_at).date() < today):
            # Check if user has an approved leave today
            approved_leave = LeaveRequest.objects.filter(
                user=u,
                status="approved",
                start_date__lte=today,
                end_date__gte=today
            ).first()
            if approved_leave:
                st.status = "on_leave"
                st.status_message = f"On Leave ({approved_leave.get_leave_type_display()})"
            else:
                att = Attendance.objects.filter(user=u, date=today).first()
                if att and att.punch_in and not att.punch_out:
                    st.status = "in_office"
                    st.status_message = ""
                else:
                    st.status = "in_office"
                    st.status_message = ""
            st.save()

    user_status.refresh_from_db()

    if request.method == "POST":
        action = request.POST.get("action", "update")
        target_id = request.POST.get("target_id")

        if target_id and (request.user.is_superuser or request.user.is_staff or str(request.user.id) == str(target_id)):
            target_user = get_object_or_404(UserModel, id=target_id)
            target_status, _ = UserStatus.objects.get_or_create(user=target_user)
        else:
            target_status = user_status

        if action == "delete":
            target_status.status_message = ""
            target_status.save()
            messages.success(request, f"Status note cleared for {target_status.user.get_full_name() or target_status.user.username}.")
        else:
            new_status = request.POST.get("status")
            msg = request.POST.get("status_message", "").strip()
            if new_status in dict(UserStatus.STATUS_CHOICES):
                target_status.status = new_status
                target_status.status_message = msg
                target_status.save()
                messages.success(request, f"Status updated for {target_status.user.get_full_name() or target_status.user.username}: {target_status.get_status_display()}")
        return redirect("status_board")

    selected_filter = request.GET.get("filter", "all").strip().lower()
    if selected_filter not in ["in_office", "remote", "meeting", "on_leave", "out_of_office"]:
        selected_filter = "all"
    search_query = request.GET.get("q", "").strip()

    if selected_date == today:
        all_statuses = UserStatus.objects.select_related("user", "user__profile").filter(user__is_active=True)
        total_team_count = all_statuses.count()
        in_office_count = all_statuses.filter(status="in_office").count()
        remote_count = all_statuses.filter(status="remote").count()
        meeting_count = all_statuses.filter(status="meeting").count()
        on_leave_count = all_statuses.filter(status="on_leave").count()
        out_of_office_count = all_statuses.filter(status="out_of_office").count()

        filtered_statuses = all_statuses
        if selected_filter != "all":
            filtered_statuses = filtered_statuses.filter(status=selected_filter)

        if search_query:
            filtered_statuses = filtered_statuses.filter(
                Q(user__username__icontains=search_query) |
                Q(user__first_name__icontains=search_query) |
                Q(user__last_name__icontains=search_query) |
                Q(user__email__icontains=search_query) |
                Q(status_message__icontains=search_query)
            )
    else:
        # Compute status for historical or future selected_date
        status_map_choices = dict(UserStatus.STATUS_CHOICES)
        all_statuses_list = []
        for u in active_users:
            leave = LeaveRequest.objects.filter(
                user=u,
                status="approved",
                start_date__lte=selected_date,
                end_date__gte=selected_date
            ).first()
            att = Attendance.objects.filter(user=u, date=selected_date).first()
            if leave:
                stat = "on_leave"
                msg = f"On Leave ({leave.get_leave_type_display()})"
            elif att and att.punch_in:
                stat = "in_office"
                msg = f"Logged {att.formatted_duration}" if att.punch_out else "Active in Office"
            elif att and att.status == "absent":
                stat = "out_of_office"
                msg = "Absent"
            else:
                stat = "in_office"
                msg = ""

            all_statuses_list.append({
                "user": u,
                "status": stat,
                "status_message": msg,
                "get_status_display": status_map_choices.get(stat, stat.replace("_", " ").title()),
            })

        total_team_count = len(all_statuses_list)
        in_office_count = sum(1 for s in all_statuses_list if s["status"] == "in_office")
        remote_count = sum(1 for s in all_statuses_list if s["status"] == "remote")
        meeting_count = sum(1 for s in all_statuses_list if s["status"] == "meeting")
        on_leave_count = sum(1 for s in all_statuses_list if s["status"] == "on_leave")
        out_of_office_count = sum(1 for s in all_statuses_list if s["status"] == "out_of_office")

        if selected_filter != "all":
            filtered_statuses = [s for s in all_statuses_list if s["status"] == selected_filter]
        else:
            filtered_statuses = all_statuses_list

        if search_query:
            q_lower = search_query.lower()
            filtered_statuses = [
                s for s in filtered_statuses
                if q_lower in s["user"].username.lower()
                or q_lower in s["user"].get_full_name().lower()
                or q_lower in (s["user"].email or "").lower()
                or q_lower in (s["status_message"] or "").lower()
            ]

    context = {
        "active_page": "status_board",
        "user_status": user_status,
        "all_statuses": filtered_statuses,
        "total_team_count": total_team_count,
        "in_office_count": in_office_count,
        "remote_count": remote_count,
        "meeting_count": meeting_count,
        "on_leave_count": on_leave_count,
        "out_of_office_count": out_of_office_count,
        "selected_filter": selected_filter,
        "search_query": search_query,
        "today_attendance": today_attendance,
        "today_date": today,
        "today_date_str": today.strftime("%Y-%m-%d"),
        "today_date_display": today.strftime("%d %B, %Y"),
        "selected_date": selected_date,
        "selected_date_str": selected_date.strftime("%Y-%m-%d"),
        "selected_date_display": selected_date.strftime("%d %B, %Y"),
        "is_today": (selected_date == today),
    }
    return render(request, "accounts/status_board.html", context)


from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger

@login_required(login_url="login")
def attendance_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    all_attendances = Attendance.objects.filter(user=request.user)

    # Active sub-tab: 'my_logs' or 'daily_team'
    active_tab = request.GET.get("tab", "").strip().lower()
    if not active_tab:
        if any(k in request.GET for k in ["daily_date", "team_status", "team_q", "team_page"]):
            active_tab = "daily_team"
        else:
            active_tab = "my_logs"
    elif active_tab not in ["my_logs", "daily_team"]:
        active_tab = "my_logs"

    # User's Leave Balances
    user_leaves = LeaveRequest.objects.filter(user=request.user)
    casual_used = sum(l.days_count for l in user_leaves.filter(leave_type="casual", status="approved"))
    sick_used = sum(l.days_count for l in user_leaves.filter(leave_type="sick", status="approved"))
    annual_used = sum(l.days_count for l in user_leaves.filter(leave_type="annual", status="approved"))
    used_leave = casual_used + sick_used + annual_used
    total_entitlement = 24
    total_available = max(0, total_entitlement - used_leave)
    pending_leave = sum(l.days_count for l in user_leaves.filter(status="pending"))

    leave_balances = {
        "total_available": total_available,
        "used_leave": used_leave,
        "pending_leave": pending_leave,
    }

    # --- Week vs Month Stats for Attendance Overview with Period Comparison & Late Tracking ---
    start_of_week = today - timedelta(days=today.weekday())
    week_qs = all_attendances.filter(date__gte=start_of_week, date__lte=today)
    week_present = week_qs.filter(status="present").count()
    week_absent = week_qs.filter(status="absent").count()
    week_leave = week_qs.filter(status__in=["leave", "lop"]).count()
    week_late = sum(1 for att in week_qs if att.punch_in and timezone.localtime(att.punch_in).time() > time(9, 30))
    
    # Working days passed this week (Mon-Sat, Sunday is Week-Off)
    week_working_days = sum(1 for i in range((today - start_of_week).days + 1) if (start_of_week + timedelta(days=i)).weekday() < 6)
    week_working_days = max(1, week_working_days)
    week_pct = min(100, round((week_present / week_working_days) * 100))

    # Previous Week comparison (6 working days Mon-Sat)
    start_of_prev_week = start_of_week - timedelta(days=7)
    end_of_prev_week = start_of_week - timedelta(days=1)
    prev_week_qs = all_attendances.filter(date__gte=start_of_prev_week, date__lte=end_of_prev_week)
    prev_week_present = prev_week_qs.filter(status="present").count()
    prev_week_pct = min(100, round((prev_week_present / 6) * 100))
    week_diff = week_pct - prev_week_pct
    week_diff_str = f"+{week_diff}%" if week_diff > 0 else (f"{week_diff}%" if week_diff < 0 else "0%")

    # Month Stats
    start_of_month = today.replace(day=1)
    month_qs = all_attendances.filter(date__gte=start_of_month, date__lte=today)
    month_present = month_qs.filter(status="present").count()
    month_absent = month_qs.filter(status="absent").count()
    month_leave = month_qs.filter(status__in=["leave", "lop"]).count()
    month_late = sum(1 for att in month_qs if att.punch_in and timezone.localtime(att.punch_in).time() > time(9, 30))
    
    # Working days in month up to today (Mon-Sat, Sunday is Week-Off)
    month_working_days = sum(1 for i in range((today - start_of_month).days + 1) if (start_of_month + timedelta(days=i)).weekday() < 6)
    month_working_days = max(1, month_working_days)
    month_pct = min(100, round((month_present / month_working_days) * 100))

    # Previous Month comparison
    last_day_prev_month = start_of_month - timedelta(days=1)
    start_of_prev_month = last_day_prev_month.replace(day=1)
    prev_month_qs = all_attendances.filter(date__gte=start_of_prev_month, date__lte=last_day_prev_month)
    prev_month_present = prev_month_qs.filter(status="present").count()
    prev_month_working = max(1, sum(1 for i in range((last_day_prev_month - start_of_prev_month).days + 1) if (start_of_prev_month + timedelta(days=i)).weekday() < 6))
    prev_month_pct = min(100, round((prev_month_present / prev_month_working) * 100))
    month_diff = month_pct - prev_month_pct
    month_diff_str = f"+{month_diff}%" if month_diff > 0 else (f"{month_diff}%" if month_diff < 0 else "0%")

    overview_stats = {
        "week": {
            "present_days": week_present,
            "absent_days": week_absent,
            "leave_days": week_leave,
            "late_days": week_late,
            "pct": week_pct,
            "prev_pct": prev_week_pct,
            "diff_str": week_diff_str,
            "diff_is_pos": (week_diff >= 0),
            "total_days": (today - start_of_week).days + 1,
        },
        "month": {
            "present_days": month_present,
            "absent_days": month_absent,
            "leave_days": month_leave,
            "late_days": month_late,
            "pct": month_pct,
            "prev_pct": prev_month_pct,
            "diff_str": month_diff_str,
            "diff_is_pos": (month_diff >= 0),
            "total_days": (today - start_of_month).days + 1,
        }
    }

    # 7-Day Visual Spline Line Graph Data (Full-Width Responsive Grid: 500x155)
    week_points = []
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    xs = [50, 118, 186, 254, 322, 390, 458]
    for i in range(7):
        d = start_of_week + timedelta(days=i)
        x = xs[i]
        att = all_attendances.filter(date=d).first()
        leave = LeaveRequest.objects.filter(user=request.user, start_date__lte=d, end_date__gte=d, status="approved").first()
        
        # Tooltip data values
        p_in_str = timezone.localtime(att.punch_in).strftime("%I:%M %p") if att and att.punch_in else "--:--"
        p_out_str = timezone.localtime(att.punch_out).strftime("%I:%M %p") if att and att.punch_out else ("Active Now" if att and att.is_punched_in else "--:--")
        dur_str = att.formatted_duration if att else "0h 0m"

        if d.weekday() == 6:
            # Sunday is always Week Off
            y = 20 if week_pct >= 80 else (78 if week_pct >= 50 else 135)
            status_type = "weekend"
            label = "Sunday (Week Off)"
            color = "#94a3b8"  # Grey
            is_non_working = True
            p_in_str = "--:--"
            p_out_str = "--:--"
            dur_str = "Week Off"
        elif d > today:
            # Future / Upcoming day
            y = 20 if week_pct >= 80 else (78 if week_pct >= 50 else 135)
            status_type = "upcoming"
            label = "Upcoming Day"
            color = "#94a3b8"  # Grey
            is_non_working = True
        else:
            # Mon-Sat Working days
            is_non_working = False
            if att and att.status == "present":
                y = 20
                status_type = "present"
                is_late = (att.punch_in and timezone.localtime(att.punch_in).time() > time(9, 30))
                label = "Present (Late Arrival)" if is_late else "Present (On-Time)"
                color = "#10b981"  # Green
            elif leave or (att and att.status in ["leave", "lop"]):
                y = 78
                status_type = "leave"
                label = f"Approved Leave ({leave.get_leave_type_display() if leave else 'PTO'})"
                color = "#f59e0b"  # Orange
            else:
                y = 135
                status_type = "absent"
                label = "Absent (Unexcused)"
                color = "#f43f5e"  # Red

        week_points.append({
            "day": day_names[i],
            "date_short": d.strftime("%b %d"),
            "date_full": d.strftime("%A, %b %d, %Y"),
            "x": x,
            "y": y,
            "status_type": status_type,
            "label": label,
            "color": color,
            "punch_in": p_in_str,
            "punch_out": p_out_str,
            "duration": dur_str,
            "is_today": (d == today),
            "is_non_working": is_non_working,
            "is_sunday": (d.weekday() == 6),
        })

    # Spline Path generation for week
    week_path_d = f"M {week_points[0]['x']},{week_points[0]['y']}"
    for i in range(len(week_points) - 1):
        p0 = week_points[i]
        p1 = week_points[i+1]
        cp1x = round(p0['x'] + (p1['x'] - p0['x']) / 2, 1)
        cp1y = p0['y']
        cp2x = round(p0['x'] + (p1['x'] - p0['x']) / 2, 1)
        cp2y = p1['y']
        week_path_d += f" C {cp1x},{cp1y} {cp2x},{cp2y} {p1['x']},{p1['y']}"
    week_area_d = f"{week_path_d} L {week_points[-1]['x']},135 L {week_points[0]['x']},135 Z"

    # Entire Month Visual Spline Line Graph Data (All Days in the Month)
    _, num_days_in_month = calendar.monthrange(today.year, today.month)
    month_points = []
    x_start = 45
    x_end = 475
    x_step = (x_end - x_start) / max(1, num_days_in_month - 1)
    
    # Calculate step interval for X-axis labels (e.g. 1st, 5th, 10th, 15th, 20th, 25th, and last day)
    label_days = {1, 5, 10, 15, 20, 25, num_days_in_month}

    for day_num in range(1, num_days_in_month + 1):
        d = start_of_month.replace(day=day_num)
        x = round(x_start + (day_num - 1) * x_step, 1)
        att = all_attendances.filter(date=d).first()
        leave = LeaveRequest.objects.filter(user=request.user, start_date__lte=d, end_date__gte=d, status="approved").first()

        p_in_str = timezone.localtime(att.punch_in).strftime("%I:%M %p") if att and att.punch_in else "--:--"
        p_out_str = timezone.localtime(att.punch_out).strftime("%I:%M %p") if att and att.punch_out else ("Active Now" if att and att.is_punched_in else "--:--")
        dur_str = att.formatted_duration if att else "0h 0m"

        if d.weekday() == 6:
            # Sunday is always Week Off
            y = 20 if month_pct >= 80 else (78 if month_pct >= 50 else 135)
            status_type = "weekend"
            label = "Sunday (Week Off)"
            color = "#94a3b8"  # Grey
            is_non_working = True
            p_in_str = "--:--"
            p_out_str = "--:--"
            dur_str = "Week Off"
        elif d > today:
            y = 20 if month_pct >= 80 else (78 if month_pct >= 50 else 135)
            status_type = "upcoming"
            label = "Upcoming Day"
            color = "#94a3b8"  # Grey
            is_non_working = True
        else:
            # Mon-Sat Working days
            is_non_working = False
            if att and att.status == "present":
                y = 20
                status_type = "present"
                is_late = (att.punch_in and timezone.localtime(att.punch_in).time() > time(9, 30))
                label = "Present (Late Arrival)" if is_late else "Present (On-Time)"
                color = "#10b981"  # Green
            elif leave or (att and att.status in ["leave", "lop"]):
                y = 78
                status_type = "leave"
                label = f"Approved Leave ({leave.get_leave_type_display() if leave else 'PTO'})"
                color = "#f59e0b"  # Orange
            else:
                y = 135
                status_type = "absent"
                label = "Absent (Unexcused)"
                color = "#f43f5e"  # Red

        month_points.append({
            "day_num": day_num,
            "day": str(day_num),
            "date_short": d.strftime("%b %d"),
            "date_full": d.strftime("%A, %b %d, %Y"),
            "x": x,
            "y": y,
            "status_type": status_type,
            "label": label,
            "color": color,
            "punch_in": p_in_str,
            "punch_out": p_out_str,
            "duration": dur_str,
            "is_today": (d == today),
            "is_non_working": is_non_working,
            "is_sunday": (d.weekday() == 6),
            "show_x_label": (day_num in label_days),
        })

    month_path_d = f"M {month_points[0]['x']},{month_points[0]['y']}"
    for i in range(len(month_points) - 1):
        p0 = month_points[i]
        p1 = month_points[i+1]
        cp1x = round(p0['x'] + (p1['x'] - p0['x']) / 2, 1)
        cp1y = p0['y']
        cp2x = round(p0['x'] + (p1['x'] - p0['x']) / 2, 1)
        cp2y = p1['y']
        month_path_d += f" C {cp1x},{cp1y} {cp2x},{cp2y} {p1['x']},{p1['y']}"
    month_area_d = f"{month_path_d} L {month_points[-1]['x']},135 L {month_points[0]['x']},135 Z"

    # --- 1. My Attendance History & Shift Logs ---
    my_attendances = all_attendances
    total_count = my_attendances.count()
    present_count = my_attendances.filter(status="present").count()
    absent_count = my_attendances.filter(status="absent").count()
    leave_count = my_attendances.filter(status="leave").count()
    lop_count = my_attendances.filter(status="lop").count()

    selected_status = request.GET.get("status", "all").strip().lower()
    if selected_status in ["present", "absent", "leave", "lop"]:
        my_attendances = my_attendances.filter(status=selected_status)
    else:
        selected_status = "all"

    search_query = request.GET.get("q", "").strip()
    if search_query:
        my_attendances = my_attendances.filter(
            Q(notes__icontains=search_query) |
            Q(date__icontains=search_query) |
            Q(status__icontains=search_query)
        )

    my_paginator = Paginator(my_attendances, 10)
    page_number = request.GET.get("page", 1)
    my_page_obj = my_paginator.get_page(page_number)

    # --- 2. Daily Team Attendance of Every Employee ---
    daily_date_str = request.GET.get("daily_date", "").strip()
    try:
        daily_date = datetime.strptime(daily_date_str, "%Y-%m-%d").date() if daily_date_str else today
    except ValueError:
        daily_date = today

    prev_daily_date = daily_date - timedelta(days=1)
    next_daily_date = daily_date + timedelta(days=1)

    all_users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("first_name", "username")
    existing_daily_atts = {att.user_id: att for att in Attendance.objects.filter(date=daily_date).select_related("user", "user__profile")}
    approved_leaves = {lr.user_id: lr for lr in LeaveRequest.objects.filter(start_date__lte=daily_date, end_date__gte=daily_date, status="approved")}
    rejected_leaves = {lr.user_id: lr for lr in LeaveRequest.objects.filter(start_date__lte=daily_date, end_date__gte=daily_date, status="rejected")}

    daily_team_records = []
    for u in all_users:
        att = existing_daily_atts.get(u.id)

        if att and att.punch_in:
            # Punched In -> Present
            stat = "present"
            p_in_formatted = timezone.localtime(att.punch_in).strftime("%I:%M %p")
            p_out_formatted = timezone.localtime(att.punch_out).strftime("%I:%M %p") if att.punch_out else ("Active Now" if att.date == today else "Auto Closed")
            duration_str = att.formatted_duration
            notes_str = att.notes or "Biometric shift recorded."
            is_active_now = (att.date == today and att.punch_out is None)
            sort_dt = att.punch_in
            priority = 1 if is_active_now else 2
        elif u.id in approved_leaves:
            # Applied Leave & Approved -> Leave
            stat = "leave"
            p_in_formatted = "--:--"
            p_out_formatted = "--:--"
            duration_str = "0h 0m"
            notes_str = f"Approved {approved_leaves[u.id].get_leave_type_display()}"
            is_active_now = False
            sort_dt = None
            priority = 3
        elif u.id in rejected_leaves:
            # Applied Leave & Rejected -> Absent
            stat = "absent"
            p_in_formatted = "--:--"
            p_out_formatted = "--:--"
            duration_str = "0h 0m"
            notes_str = "Leave rejected - Marked Absent"
            is_active_now = False
            sort_dt = None
            priority = 4
        else:
            # Not punched in & no approved leave -> Starts day as Absent
            stat = "absent"
            p_in_formatted = "--:--"
            p_out_formatted = "--:--"
            duration_str = "0h 0m"
            notes_str = "Not punched in - Marked Absent"
            is_active_now = False
            sort_dt = None
            priority = 4

        u_name = u.get_full_name() or u.username
        u_init = "".join([part[0] for part in u_name.split()][:2]).upper() or "EM"
        avatar_url = u.profile.avatar.url if (hasattr(u, 'profile') and u.profile.avatar) else None
        dept = u.profile.department if (hasattr(u, 'profile') and u.profile.department) else "Engineering"
        desig = u.profile.designation if (hasattr(u, 'profile') and u.profile.designation) else "Staff"

        daily_team_records.append({
            "user_id": u.id,
            "name": u_name,
            "username": u.username,
            "initials": u_init,
            "avatar_url": avatar_url,
            "department": dept,
            "designation": desig,
            "punch_in": p_in_formatted,
            "punch_out": p_out_formatted,
            "is_active_now": is_active_now,
            "duration": duration_str,
            "status": stat,
            "status_display": "Present" if stat == "present" else ("Absent" if stat == "absent" else ("Leave" if stat == "leave" else "LOP")),
            "notes": notes_str,
            "priority": priority,
            "sort_dt": sort_dt,
        })

    # Sort with time (descending): Active check-ins first (latest punch_in first), then completed shifts, then leaves, then absents
    def record_sort_key(item):
        ts = item["sort_dt"].timestamp() if item["sort_dt"] else 0
        return (item["priority"], -ts, item["name"])

    daily_team_records.sort(key=record_sort_key)

    # Summary metrics for Daily Team Attendance
    team_total_count = len(daily_team_records)
    team_present_count = sum(1 for r in daily_team_records if r["status"] == "present")
    team_absent_count = sum(1 for r in daily_team_records if r["status"] == "absent")
    team_leave_count = sum(1 for r in daily_team_records if r["status"] == "leave")
    team_lop_count = sum(1 for r in daily_team_records if r["status"] == "lop")

    # Filter Daily Team records by status & search
    team_status_filter = request.GET.get("team_status", "all").strip().lower()
    if team_status_filter in ["present", "absent", "leave", "lop"]:
        filtered_team_records = [r for r in daily_team_records if r["status"] == team_status_filter]
    else:
        team_status_filter = "all"
        filtered_team_records = daily_team_records

    team_search_query = request.GET.get("team_q", "").strip().lower()
    if team_search_query:
        filtered_team_records = [
            r for r in filtered_team_records
            if team_search_query in r["name"].lower()
            or team_search_query in r["username"].lower()
            or team_search_query in r["department"].lower()
            or team_search_query in r["designation"].lower()
        ]

    # Paginate Daily Team Records: 10 per page
    team_paginator = Paginator(filtered_team_records, 10)
    team_page_number = request.GET.get("team_page", 1)
    team_page_obj = team_paginator.get_page(team_page_number)

    context = {
        "active_page": "attendance",
        "active_tab": active_tab,
        "today": today,
        "today_attendance": today_attendance,
        # Leave Balances
        "leave_balances": leave_balances,
        # Overview Stats
        "overview_stats": overview_stats,
        "week_stats": overview_stats["week"],
        "month_stats": overview_stats["month"],
        "week_points": week_points,
        "week_path_d": week_path_d,
        "week_area_d": week_area_d,
        "month_points": month_points,
        "month_path_d": month_path_d,
        "month_area_d": month_area_d,
        # Personal Logs
        "page_obj": my_page_obj,
        "attendances": my_page_obj.object_list,
        "selected_status": selected_status,
        "search_query": search_query,
        "total_count": total_count,
        "present_count": present_count,
        "absent_count": absent_count,
        "leave_count": leave_count,
        "lop_count": lop_count,
        # Daily Team Attendance
        "daily_date": daily_date,
        "daily_date_str": daily_date.strftime("%Y-%m-%d"),
        "prev_daily_date_str": prev_daily_date.strftime("%Y-%m-%d"),
        "next_daily_date_str": next_daily_date.strftime("%Y-%m-%d"),
        "is_today_daily": (daily_date == today),
        "team_page_obj": team_page_obj,
        "team_records": team_page_obj.object_list,
        "team_status_filter": team_status_filter,
        "team_search_query": team_search_query,
        "team_total_count": team_total_count,
        "team_present_count": team_present_count,
        "team_absent_count": team_absent_count,
        "team_leave_count": team_leave_count,
        "team_lop_count": team_lop_count,
    }
    return render(request, "accounts/attendance.html", context)


@login_required(login_url="login")
def punch_attendance_view(request):
    if request.method == "POST":
        action = request.POST.get("action")
        today = timezone.localdate()
        now = timezone.now()

        # Auto-close any lingering unclosed shifts from previous calendar days prior to 12:00 AM
        prev_open_att = Attendance.objects.filter(
            user=request.user,
            date__lt=today,
            punch_in__isnull=False,
            punch_out__isnull=True,
        ).first()
        if prev_open_att:
            prev_end = timezone.make_aware(datetime.combine(prev_open_att.date, time(23, 59, 59)))
            prev_open_att.punch_out = prev_end
            prev_open_att.save()

        # Get or create attendance for the current day (after 12:00 AM considered new day)
        attendance, created = Attendance.objects.get_or_create(
            user=request.user,
            date=today,
            defaults={"status": "present", "notes": "Biometric portal entry"},
        )

        if action == "punch_in":
            if not attendance.punch_in:
                attendance.punch_in = now
                attendance.status = "present"
                attendance.notes = "Checked in via Portal"
                attendance.save()

                # Update live presence status
                user_status, _ = UserStatus.objects.get_or_create(user=request.user)
                user_status.status = "in_office"
                user_status.save()

                messages.success(request, f"Checked in successfully at {timezone.localtime(now).strftime('%I:%M %p')}. Attendance marked as Present for {today.strftime('%d %B %Y')}.")
            else:
                messages.warning(request, "You have already checked in today.")

        elif action == "punch_out":
            if attendance.punch_in and not attendance.punch_out:
                attendance.punch_out = now
                attendance.status = "present"
                attendance.save()

                # Update live presence status
                user_status, _ = UserStatus.objects.get_or_create(user=request.user)
                if user_status.status == "in_office":
                    user_status.status = "out_of_office"
                    user_status.save()

                messages.success(request, f"Checked out successfully at {timezone.localtime(now).strftime('%I:%M %p')}. Total shift duration: {attendance.formatted_duration}.")
            elif not attendance.punch_in:
                messages.error(request, "You need to check in before checking out.")
            else:
                messages.warning(request, "You have already checked out today.")

    next_url = request.META.get("HTTP_REFERER") or "home"
    return redirect(next_url)


@login_required(login_url="login")
def tasks_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    
    # Base task queryset
    base_tasks = DailyTask.objects.select_related("user", "user__profile").all()

    total_tasks = base_tasks.count()
    pending_count = base_tasks.filter(status="pending").count()
    in_progress_count = base_tasks.filter(status="in_progress").count()
    completed_count = base_tasks.filter(status="completed").count()
    completion_rate = int((completed_count / total_tasks * 100)) if total_tasks > 0 else 0

    # Status Filtering
    selected_status = request.GET.get("status", "all").strip().lower()
    filtered_qs = base_tasks

    if selected_status in ["pending", "in_progress", "completed"]:
        filtered_qs = filtered_qs.filter(status=selected_status)
    else:
        selected_status = "all"

    # Search Query Filtering
    search_query = request.GET.get("q", "").strip()
    if search_query:
        filtered_qs = filtered_qs.filter(
            Q(title__icontains=search_query) |
            Q(description__icontains=search_query) |
            Q(project__icontains=search_query) |
            Q(task_type__icontains=search_query) |
            Q(user__first_name__icontains=search_query) |
            Q(user__last_name__icontains=search_query) |
            Q(user__username__icontains=search_query) |
            Q(assigned_by__icontains=search_query)
        )

    # Date Range Filtering (From Date - To Date)
    start_date_str = request.GET.get("start_date", "").strip()
    end_date_str = request.GET.get("end_date", "").strip()

    if start_date_str:
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            filtered_qs = filtered_qs.filter(Q(due_date__gte=start_date) | Q(date__gte=start_date))
        except ValueError:
            pass

    if end_date_str:
        try:
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
            filtered_qs = filtered_qs.filter(Q(due_date__lte=end_date) | Q(date__lte=end_date))
        except ValueError:
            pass

    # Order by newest
    filtered_qs = filtered_qs.order_by("-id")

    # Pagination: 8 tasks per page
    paginator = Paginator(filtered_qs, 8)
    page_number = request.GET.get("page", 1)
    page_obj = paginator.get_page(page_number)

    all_users = UserModel.objects.filter(is_active=True).order_by("first_name", "username")
    task_form = TaskForm()

    context = {
        "active_page": "tasks",
        "search_query": search_query,
        "selected_status": selected_status,
        "start_date": start_date_str,
        "end_date": end_date_str,
        "page_obj": page_obj,
        "tasks": page_obj.object_list,
        "total_tasks": total_tasks,
        "pending_count": pending_count,
        "in_progress_count": in_progress_count,
        "completed_count": completed_count,
        "completion_rate": completion_rate,
        "task_form": task_form,
        "all_users": all_users,
        "today_attendance": today_attendance,
    }
    return render(request, "accounts/tasks.html", context)


@login_required(login_url="login")
def task_create_view(request):
    if request.method == "POST":
        title = request.POST.get("title", "").strip()
        description = request.POST.get("description", "").strip()
        task_type = request.POST.get("task_type", "dev").strip()
        project = request.POST.get("project", "HRMS Portal").strip() or "HRMS Portal"
        priority = request.POST.get("priority", "medium").strip()
        due_date_str = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", "pending").strip()
        assigned_by = request.POST.get("assigned_by", "").strip()
        user_id = request.POST.get("user_id", "").strip()

        if not assigned_by:
            assigned_by = request.user.get_full_name() or request.user.username

        assignee = request.user
        if user_id:
            try:
                assignee = UserModel.objects.get(pk=user_id)
            except UserModel.DoesNotExist:
                assignee = request.user

        due_date = None
        if due_date_str:
            try:
                due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
            except ValueError:
                due_date = None

        if title:
            DailyTask.objects.create(
                user=assignee,
                title=title,
                description=description,
                task_type=task_type,
                project=project,
                assigned_by=assigned_by,
                priority=priority,
                due_date=due_date,
                status=status,
            )
            messages.success(request, f"Task '{title}' created successfully!")
        else:
            messages.error(request, "Task title cannot be empty.")

    next_url = request.META.get("HTTP_REFERER") or "tasks"
    return redirect(next_url)


@login_required(login_url="login")
def task_edit_view(request, pk):
    task = get_object_or_404(DailyTask, pk=pk)
    if request.method == "POST":
        title = request.POST.get("title", "").strip()
        description = request.POST.get("description", "").strip()
        task_type = request.POST.get("task_type", task.task_type).strip()
        project = request.POST.get("project", task.project).strip()
        priority = request.POST.get("priority", task.priority).strip()
        due_date_str = request.POST.get("due_date", "").strip()
        status = request.POST.get("status", task.status).strip()
        assigned_by = request.POST.get("assigned_by", task.assigned_by).strip()
        user_id = request.POST.get("user_id", "").strip()

        if user_id:
            try:
                task.user = UserModel.objects.get(pk=user_id)
            except UserModel.DoesNotExist:
                pass

        if title:
            task.title = title
            task.description = description
            task.task_type = task_type
            task.project = project
            task.priority = priority
            task.status = status
            task.assigned_by = assigned_by

            if due_date_str:
                try:
                    task.due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                except ValueError:
                    task.due_date = None
            else:
                task.due_date = None

            task.save()
            messages.success(request, f"Task '{task.title}' updated successfully!")
        else:
            messages.error(request, "Task title cannot be empty.")

    next_url = request.META.get("HTTP_REFERER") or "tasks"
    return redirect(next_url)


@login_required(login_url="login")
def task_update_status_view(request, pk):
    task = get_object_or_404(DailyTask, pk=pk)
    if request.method == "POST":
        new_status = request.POST.get("status")
        if new_status in dict(DailyTask.STATUS_CHOICES):
            task.status = new_status
            task.save()
            messages.success(request, f"Task '{task.title}' marked as {task.get_status_display()}.")
    next_url = request.META.get("HTTP_REFERER") or "tasks"
    return redirect(next_url)


@login_required(login_url="login")
def task_delete_view(request, pk):
    task = get_object_or_404(DailyTask, pk=pk)
    if request.method == "POST":
        task_title = task.title
        task.delete()
        messages.info(request, f"Task '{task_title}' deleted.")
    next_url = request.META.get("HTTP_REFERER") or "tasks"
    return redirect(next_url)


@login_required(login_url="login")
def tasks_export_excel_view(request):
    base_tasks = DailyTask.objects.select_related("user", "user__profile").all()

    selected_status = request.GET.get("status", "all").strip().lower()
    if selected_status in ["pending", "in_progress", "completed"]:
        base_tasks = base_tasks.filter(status=selected_status)

    search_query = request.GET.get("q", "").strip()
    if search_query:
        base_tasks = base_tasks.filter(
            Q(title__icontains=search_query) |
            Q(description__icontains=search_query) |
            Q(project__icontains=search_query) |
            Q(task_type__icontains=search_query) |
            Q(user__first_name__icontains=search_query) |
            Q(user__last_name__icontains=search_query) |
            Q(user__username__icontains=search_query) |
            Q(assigned_by__icontains=search_query)
        )

    start_date_str = request.GET.get("start_date", "").strip()
    end_date_str = request.GET.get("end_date", "").strip()
    if start_date_str:
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            base_tasks = base_tasks.filter(Q(due_date__gte=start_date) | Q(date__gte=start_date))
        except ValueError:
            pass
    if end_date_str:
        try:
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
            base_tasks = base_tasks.filter(Q(due_date__lte=end_date) | Q(date__lte=end_date))
        except ValueError:
            pass

    base_tasks = base_tasks.order_by("-id")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Tasks"
    ws.views.sheetView[0].showGridLines = True

    # Styling
    title_font = Font(name="Arial", size=14, bold=True, color="7C5DFA")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=9, color="1E293B")
    header_fill = PatternFill(start_color="7C5DFA", end_color="7C5DFA", fill_type="solid")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    ws.merge_cells("A1:J1")
    ws["A1"] = f"Daily Tasks Roster - Exported on {timezone.localdate().strftime('%d %B %Y')}"
    ws["A1"].font = title_font

    headers = [
        "Task ID", "Type", "Assignee Name", "Emp ID", "Project",
        "Task Title", "Priority", "Due Date", "Assigned By", "Status"
    ]
    ws.append([]) # row 2
    ws.append(headers) # row 3

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, task in enumerate(base_tasks, start=4):
        emp_name = task.user.get_full_name() or task.user.username
        emp_id = f"EMP-{task.user.id:04d}"
        due_date_formatted = task.due_date.strftime("%d %b, %Y") if task.due_date else "--"
        row_data = [
            f"TSK-{task.id:04d}",
            task.get_task_type_display() if hasattr(task, 'get_task_type_display') else (task.task_type or "Dev"),
            emp_name,
            emp_id,
            task.project or "HRMS Portal",
            task.title,
            task.get_priority_display(),
            due_date_formatted,
            task.assigned_by or "Admin",
            task.get_status_display(),
        ]
        ws.append(row_data)
        fill = alt_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.font = data_font
            c.border = thin_border
            if fill.fill_type:
                c.fill = fill

    col_widths = [12, 16, 22, 14, 18, 34, 12, 14, 20, 14]
    for i, w in enumerate(col_widths, start=1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = w

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    filename = f"Daily_Tasks_Export_{timezone.localdate().strftime('%Y%m%d')}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


@login_required(login_url="login")
def leaves_view(request):
    if request.user.is_superuser or request.user.is_staff:
        base_leaves = LeaveRequest.objects.select_related("user", "user__profile").all()
    else:
        base_leaves = LeaveRequest.objects.filter(user=request.user)

    user_leaves = LeaveRequest.objects.filter(user=request.user)
    form = LeaveRequestForm()
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()

    # Leave balances (defaults for standard HR policy)
    casual_used = sum(l.days_count for l in user_leaves.filter(leave_type="casual", status="approved"))
    sick_used = sum(l.days_count for l in user_leaves.filter(leave_type="sick", status="approved"))
    annual_used = sum(l.days_count for l in user_leaves.filter(leave_type="annual", status="approved"))

    leave_balances = {
        "casual": max(0, 12 - casual_used),
        "sick": max(0, 8 - sick_used),
        "annual": max(0, 15 - annual_used),
        "total_remaining": (12 - casual_used) + (8 - sick_used) + (15 - annual_used),
    }

    # Tab and Search Filters
    selected_tab = request.GET.get("tab", "").strip().lower()
    status_filter = request.GET.get("status", "").strip().lower()
    if not selected_tab and status_filter:
        selected_tab = status_filter
    if not selected_tab:
        selected_tab = "all"

    search_query = request.GET.get("q", "").strip()

    if selected_tab == "my":
        filtered_leaves = base_leaves.filter(user=request.user)
    elif selected_tab == "pending":
        filtered_leaves = base_leaves.filter(status="pending")
    elif selected_tab == "approved":
        filtered_leaves = base_leaves.filter(status="approved")
    elif selected_tab == "rejected":
        filtered_leaves = base_leaves.filter(status="rejected")
    else:
        filtered_leaves = base_leaves

    if search_query:
        filtered_leaves = filtered_leaves.filter(
            Q(user__first_name__icontains=search_query) |
            Q(user__last_name__icontains=search_query) |
            Q(user__username__icontains=search_query) |
            Q(reason__icontains=search_query) |
            Q(leave_type__icontains=search_query)
        )

    # Order by creation/date desc
    filtered_leaves = filtered_leaves.order_by("-id")

    # Pagination: 8 per page
    paginator = Paginator(filtered_leaves, 8)
    page_number = request.GET.get("page", 1)
    page_obj = paginator.get_page(page_number)

    context = {
        "active_page": "leaves",
        "leaves": page_obj,
        "page_obj": page_obj,
        "paginator": paginator,
        "form": form,
        "leave_balances": leave_balances,
        "today_attendance": today_attendance,
        "selected_tab": selected_tab,
        "selected_status": selected_tab,
        "search_query": search_query,
        "total_leaves_count": base_leaves.count(),
        "total_count": base_leaves.count(),
        "my_total": base_leaves.filter(user=request.user).count(),
        "pending_total": base_leaves.filter(status="pending").count(),
        "pending_count": base_leaves.filter(status="pending").count(),
        "approved_total": base_leaves.filter(status="approved").count(),
        "approved_count": base_leaves.filter(status="approved").count(),
        "rejected_total": base_leaves.filter(status="rejected").count(),
        "rejected_count": base_leaves.filter(status="rejected").count(),
    }
    return render(request, "accounts/leaves.html", context)


@login_required(login_url="login")
def leave_apply_view(request):
    if request.method == "POST":
        form = LeaveRequestForm(request.POST)
        if form.is_valid():
            leave = form.save(commit=False)
            leave.user = request.user
            leave.save()
            messages.success(request, "Leave request submitted successfully!")
        else:
            messages.error(request, "Please check the dates and details.")
        next_url = request.POST.get("next") or request.META.get("HTTP_REFERER") or "leaves"
        return redirect(next_url)
    return redirect("leaves")


@login_required(login_url="login")
def leave_edit_view(request, pk):
    if request.user.is_superuser or request.user.is_staff:
        leave = get_object_or_404(LeaveRequest, pk=pk)
    else:
        leave = get_object_or_404(LeaveRequest, pk=pk, user=request.user)

    if request.method == "POST":
        form = LeaveRequestForm(request.POST, instance=leave)
        if form.is_valid():
            form.save()
            messages.success(request, "Leave application updated successfully!")
        else:
            messages.error(request, "Failed to update leave application.")
    return redirect("leaves")


@login_required(login_url="login")
def leave_delete_view(request, pk):
    if request.user.is_superuser or request.user.is_staff:
        leave = get_object_or_404(LeaveRequest, pk=pk)
    else:
        leave = get_object_or_404(LeaveRequest, pk=pk, user=request.user)

    if request.method == "POST":
        leave.delete()
        messages.info(request, "Leave application cancelled/removed.")
    return redirect("leaves")


@login_required(login_url="login")
def leave_status_update_view(request, pk):
    leave = get_object_or_404(LeaveRequest, pk=pk)
    if request.method == "POST":
        new_status = request.POST.get("status")
        if new_status in ["approved", "rejected", "pending"]:
            leave.status = new_status
            leave.save()

            # Synchronize daily attendance records for the leave date range
            cur_date = leave.start_date
            while cur_date <= leave.end_date:
                att = Attendance.objects.filter(user=leave.user, date=cur_date).first()
                if new_status == "approved":
                    if not att or not att.punch_in:
                        Attendance.objects.update_or_create(
                            user=leave.user,
                            date=cur_date,
                            defaults={
                                "status": "leave",
                                "punch_in": None,
                                "punch_out": None,
                                "notes": f"Approved {leave.get_leave_type_display()}",
                            },
                        )
                elif new_status == "rejected":
                    if att and not att.punch_in:
                        att.status = "absent"
                        att.notes = "Leave rejected - Marked Absent"
                        att.save()
                cur_date += timedelta(days=1)

            msg = f"Leave request for {leave.user.get_full_name() or leave.user.username} marked as {leave.get_status_display()}."
            messages.success(request, msg)
            if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.POST.get("ajax") == "1":
                return JsonResponse({
                    "status": "success",
                    "new_status": new_status,
                    "status_display": leave.get_status_display(),
                    "message": msg,
                })
    next_url = request.META.get("HTTP_REFERER") or "home"
    return redirect(next_url)


@login_required(login_url="login")
def permissions_view(request):
    """Full-featured permissions and gate pass management dashboard."""
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    form = PermissionRequestForm()

    # Base QuerySet
    if request.user.is_superuser or request.user.is_staff:
        all_permissions = PermissionRequest.objects.select_related("user", "user__profile").all()
    else:
        all_permissions = PermissionRequest.objects.filter(user=request.user)

    # Current user's monthly balance calculation
    user_month_permissions = PermissionRequest.objects.filter(
        user=request.user,
        date__year=today.year,
        date__month=today.month,
    )
    used_hours = sum(float(p.duration_hours) for p in user_month_permissions.filter(status="approved"))
    monthly_allowance = 4.0  # Standard 4.0 hours permission quota per month
    available_hours = max(0.0, monthly_allowance - used_hours)

    permission_balances = {
        "monthly_allowance": monthly_allowance,
        "used_hours": round(used_hours, 1),
        "available_hours": round(available_hours, 1),
        "approved_count": user_month_permissions.filter(status="approved").count(),
        "pending_count": user_month_permissions.filter(status="pending").count(),
        "on_duty_count": user_month_permissions.filter(permission_type="on_duty").count(),
    }

    # Tab and Search Filters
    selected_filter = request.GET.get("tab", "all").strip().lower()
    search_query = request.GET.get("q", "").strip()

    if selected_filter == "my":
        filtered_permissions = all_permissions.filter(user=request.user)
    elif selected_filter == "pending":
        filtered_permissions = all_permissions.filter(status="pending")
    elif selected_filter == "approved":
        filtered_permissions = all_permissions.filter(status="approved")
    elif selected_filter == "rejected":
        filtered_permissions = all_permissions.filter(status="rejected")
    elif selected_filter == "on_duty":
        filtered_permissions = all_permissions.filter(permission_type="on_duty")
    else:
        filtered_permissions = all_permissions

    if search_query:
        filtered_permissions = filtered_permissions.filter(
            Q(user__first_name__icontains=search_query) |
            Q(user__last_name__icontains=search_query) |
            Q(user__username__icontains=search_query) |
            Q(reason__icontains=search_query) |
            Q(permission_type__icontains=search_query)
        )

    # Order by id desc (latest first)
    filtered_permissions = filtered_permissions.order_by("-id")

    # Pagination: 8 requests per page
    paginator = Paginator(filtered_permissions, 8)
    page_number = request.GET.get("page", 1)
    page_obj = paginator.get_page(page_number)

    context = {
        "active_page": "app_permissions",
        "permissions": page_obj,
        "page_obj": page_obj,
        "paginator": paginator,
        "form": form,
        "permission_balances": permission_balances,
        "today_attendance": today_attendance,
        "selected_tab": selected_filter,
        "search_query": search_query,
        "total_requests_count": all_permissions.count(),
        "pending_total": all_permissions.filter(status="pending").count(),
        "approved_total": all_permissions.filter(status="approved").count(),
        "rejected_total": all_permissions.filter(status="rejected").count(),
    }
    return render(request, "accounts/permissions.html", context)


@login_required(login_url="login")
def permission_apply_view(request):
    if request.method == "POST":
        form = PermissionRequestForm(request.POST)
        if form.is_valid():
            perm = form.save(commit=False)
            perm.user = request.user
            perm.save()
            messages.success(request, "Permission / Gate Pass request submitted successfully!")
        else:
            messages.error(request, "Please verify all required fields for permission request.")
    return redirect("permissions")


@login_required(login_url="login")
def permission_edit_view(request, pk):
    if request.user.is_superuser or request.user.is_staff:
        perm = get_object_or_404(PermissionRequest, pk=pk)
    else:
        perm = get_object_or_404(PermissionRequest, pk=pk, user=request.user)

    if request.method == "POST":
        form = PermissionRequestForm(request.POST, instance=perm)
        if form.is_valid():
            form.save()
            messages.success(request, "Permission request updated successfully!")
        else:
            messages.error(request, "Failed to update permission request.")
    return redirect("permissions")


@login_required(login_url="login")
def permission_delete_view(request, pk):
    if request.user.is_superuser or request.user.is_staff:
        perm = get_object_or_404(PermissionRequest, pk=pk)
    else:
        perm = get_object_or_404(PermissionRequest, pk=pk, user=request.user)

    if request.method == "POST":
        perm.delete()
        messages.info(request, "Permission request removed.")
    return redirect("permissions")


@login_required(login_url="login")
def permission_status_update_view(request, pk):
    perm = get_object_or_404(PermissionRequest, pk=pk)
    if request.method == "POST":
        new_status = request.POST.get("status")
        if new_status in ["approved", "rejected", "pending"]:
            perm.status = new_status
            perm.save()
            msg = f"Permission for {perm.user.get_full_name() or perm.user.username} marked as {perm.get_status_display()}."
            messages.success(request, msg)
            if request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.POST.get("ajax") == "1":
                return JsonResponse({
                    "status": "success",
                    "new_status": new_status,
                    "status_display": perm.get_status_display(),
                    "message": msg,
                })
    next_url = request.META.get("HTTP_REFERER") or "permissions"
    return redirect(next_url)


import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from django.http import HttpResponse


DEV_USERNAMES = {
    "alex.chen", "priya.sharma", "rahul.varma", "elena.rostova",
    "siddharth.patel", "clara.zhao", "tariq.mansoor"
}


def _get_user_role_details(emp):
    if emp.is_superuser:
        return "Super Admin", "badge-superadmin", ""
    elif emp.is_staff:
        return "HR / Staff Admin", "badge-staff", ""
    elif emp.username in DEV_USERNAMES or "dev" in emp.username.lower():
        return "Developer", "badge-developer", ""
    return "Employee", "badge-employee", ""


def _get_filtered_employees(request):
    search_query = request.GET.get("q", "").strip()
    role_filter = request.GET.get("role", "all").strip().lower()
    status_filter = request.GET.get("status", "all").strip().lower()
    active_filter = request.GET.get("active", "all").strip().lower()

    queryset = UserModel.objects.select_related("work_status").order_by("id")

    if search_query:
        queryset = queryset.filter(
            Q(username__icontains=search_query) |
            Q(first_name__icontains=search_query) |
            Q(last_name__icontains=search_query) |
            Q(email__icontains=search_query) |
            Q(work_status__status_message__icontains=search_query)
        )

    if role_filter == "superadmin":
        queryset = queryset.filter(is_superuser=True)
    elif role_filter == "staff":
        queryset = queryset.filter(is_staff=True, is_superuser=False)
    elif role_filter == "developer":
        queryset = queryset.filter(Q(username__in=DEV_USERNAMES) | Q(username__icontains="dev"))
    elif role_filter == "employee":
        queryset = queryset.filter(is_staff=False, is_superuser=False).exclude(username__in=DEV_USERNAMES)
    elif role_filter == "member":
        queryset = queryset.filter(is_staff=False, is_superuser=False)

    if status_filter in ["in_office", "remote", "meeting", "on_leave", "out_of_office"]:
        queryset = queryset.filter(work_status__status=status_filter)

    if active_filter == "active":
        queryset = queryset.filter(is_active=True)
    elif active_filter == "inactive":
        queryset = queryset.filter(is_active=False)

    return queryset, search_query, role_filter, status_filter, active_filter


@login_required(login_url="login")
def employees_directory_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()

    # KPI Metrics
    total_count = UserModel.objects.count()
    superadmin_count = UserModel.objects.filter(is_superuser=True).count()
    staff_count = UserModel.objects.filter(is_staff=True, is_superuser=False).count()
    dev_count = UserModel.objects.filter(Q(username__in=DEV_USERNAMES) | Q(username__icontains="dev")).count()
    members_count = UserModel.objects.filter(is_staff=False, is_superuser=False).count()
    active_count = UserModel.objects.filter(is_active=True).count()
    in_office_count = UserStatus.objects.filter(status="in_office").count()
    remote_count = UserStatus.objects.filter(status="remote").count()
    on_leave_count = UserStatus.objects.filter(status="on_leave").count()

    queryset, search_query, role_filter, status_filter, active_filter = _get_filtered_employees(request)

    # Today's attendance mapping for quick display
    today_attendances = {
        att.user_id: att for att in Attendance.objects.filter(date=today).select_related("user")
    }

    # Attach computed fields for template
    employees_list = list(queryset)
    for emp in employees_list:
        att = today_attendances.get(emp.id)
        emp.today_att = att
        emp.check_in_time = att.punch_in.strftime("%I:%M %p") if (att and att.punch_in) else None
        emp.check_out_time = att.punch_out.strftime("%I:%M %p") if (att and att.punch_out) else None
        emp.is_punched_in = att.is_punched_in if att else False
        emp.is_on_leave = (att and att.status == "leave") or (hasattr(emp, "work_status") and emp.work_status.status == "on_leave")

        role_label, role_badge, role_icon = _get_user_role_details(emp)
        emp.role_label = role_label
        emp.role_badge = role_badge
        emp.role_icon = role_icon

    # Pagination: 8 employees per page
    page_size = 8
    paginator = Paginator(employees_list, page_size)
    page_number = request.GET.get("page", 1)
    page_obj = paginator.get_page(page_number)

    context = {
        "active_page": "employees",
        "today_attendance": today_attendance,
        "page_obj": page_obj,
        "paginator": paginator,
        "total_count": total_count,
        "superadmin_count": superadmin_count,
        "staff_count": staff_count,
        "dev_count": dev_count,
        "members_count": members_count,
        "active_count": active_count,
        "in_office_count": in_office_count,
        "remote_count": remote_count,
        "on_leave_count": on_leave_count,
        "filtered_count": len(employees_list),
        "search_query": search_query,
        "role_filter": role_filter,
        "status_filter": status_filter,
        "active_filter": active_filter,
    }
    return render(request, "accounts/employees.html", context)


@login_required(login_url="login")
def teams_view(request):
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=request.user, date=today).first()
    today_attendances = {att.user_id: att for att in Attendance.objects.filter(date=today)}

    # Defined 8 Teams Specification
    teams_config = [
        {
            "key": "super_admin",
            "name": "Super Admin",
            "short_code": "SA",
            "color": "#7c5dfa",
            "bg_color": "var(--pill-purple)",
            "border_color": "rgba(124, 93, 250, 0.35)",
            "description": "Executive Leadership & Board of Directors with full system authority.",
            "members": [],
        },
        {
            "key": "admin",
            "name": "Admin",
            "short_code": "ADM",
            "color": "#3b82f6",
            "bg_color": "var(--pill-blue)",
            "border_color": "rgba(59, 130, 246, 0.35)",
            "description": "System Operations, Security Administration & IT Infrastructure management.",
            "members": [],
        },
        {
            "key": "hr_manager",
            "name": "HR Manager",
            "short_code": "HRM",
            "color": "#ec4899",
            "bg_color": "var(--pill-pink)",
            "border_color": "rgba(236, 72, 153, 0.35)",
            "description": "Strategic human resource policies, talent governance and department oversight.",
            "members": [],
        },
        {
            "key": "hr_executive",
            "name": "HR Executive",
            "short_code": "HRE",
            "color": "#8b5cf6",
            "bg_color": "rgba(139, 92, 246, 0.12)",
            "border_color": "rgba(139, 92, 246, 0.35)",
            "description": "Employee onboarding, leave administration, payroll operations and recruitment.",
            "members": [],
        },
        {
            "key": "manager",
            "name": "Manager",
            "short_code": "MGR",
            "color": "#f59e0b",
            "bg_color": "var(--pill-orange)",
            "border_color": "rgba(245, 158, 11, 0.35)",
            "description": "Engineering, Design & Operations leadership driving team sprint execution.",
            "members": [],
        },
        {
            "key": "development",
            "name": "Development",
            "short_code": "DEV",
            "color": "#10b981",
            "bg_color": "var(--pill-green)",
            "border_color": "rgba(16, 185, 129, 0.35)",
            "description": "Full-stack software engineers, mobile developers, DevOps and QA specialists.",
            "members": [],
        },
        {
            "key": "finance",
            "name": "Finance",
            "short_code": "FIN",
            "color": "#06b6d4",
            "bg_color": "var(--pill-cyan)",
            "border_color": "rgba(6, 182, 212, 0.35)",
            "description": "Financial planning, budgets, corporate accounting and tax compliance.",
            "members": [],
        },
        {
            "key": "support",
            "name": "Support",
            "short_code": "SUP",
            "color": "#6366f1",
            "bg_color": "rgba(99, 102, 241, 0.12)",
            "border_color": "rgba(99, 102, 241, 0.35)",
            "description": "Customer technical support, client success, and helpdesk operations.",
            "members": [],
        },
    ]

    teams_map = {t["key"]: t for t in teams_config}

    # Fetch all active users
    all_users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("first_name", "username")

    search_query = request.GET.get("q", "").strip().lower()
    selected_team = request.GET.get("team", "all").strip().lower()
    if selected_team not in teams_map and selected_team != "all":
        selected_team = "all"

    total_in_office = 0
    total_remote = 0
    total_on_leave = 0

    for user in all_users:
        profile = getattr(user, "profile", None)
        work_status = getattr(user, "work_status", None)
        att = today_attendances.get(user.id)

        # Team determination
        team_key = getattr(profile, "team", None) if profile else None
        if not team_key or team_key not in teams_map:
            # Fallback auto-categorization
            if user.is_superuser:
                team_key = "super_admin"
            elif user.is_staff and user.username in ["david.miller", "admin"]:
                team_key = "admin"
            elif user.username in ["marcus.vance"] or (profile and "hr operations director" in profile.designation.lower()):
                team_key = "hr_manager"
            elif user.username in ["anita.deshmukh"] or (profile and "hr executive" in profile.designation.lower()):
                team_key = "hr_executive"
            elif user.username in ["sophie.laurent", "vikram.singh"] or (profile and "manager" in profile.designation.lower()):
                team_key = "manager"
            elif user.username in ["rohan.mehra", "rachel.green"] or (profile and "finance" in profile.department.lower()):
                team_key = "finance"
            elif user.username in ["hannah.schmidt", "maya.lin"] or (profile and "support" in profile.department.lower()):
                team_key = "support"
            else:
                team_key = "development"

        # Punch / Presence status
        presence = work_status.status if work_status else "in_office"
        presence_label = work_status.get_status_display() if work_status else "In Office"
        status_msg = work_status.status_message if work_status else ""

        if presence == "in_office":
            total_in_office += 1
        elif presence in ["remote", "meeting"]:
            total_remote += 1
        elif presence in ["on_leave", "out_of_office"]:
            total_on_leave += 1

        is_punched_in = att.is_punched_in if att else False
        punch_in_str = att.punch_in.strftime("%I:%M %p") if (att and att.punch_in) else None
        punch_out_str = att.punch_out.strftime("%I:%M %p") if (att and att.punch_out) else None

        member_data = {
            "user": user,
            "id": user.id,
            "username": user.username,
            "full_name": user.get_full_name() or user.username,
            "initials": (user.first_name[:1] + user.last_name[:1]).upper() if user.first_name and user.last_name else user.username[:2].upper(),
            "email": user.email or "",
            "phone": profile.phone if profile else "+91 98765 43210",
            "department": profile.department if profile else "Engineering",
            "designation": profile.designation if profile else "Team Member",
            "work_location": profile.work_location if profile else "Bangalore HQ",
            "reporting_manager": profile.reporting_manager if profile else "Direct Lead",
            "avatar": profile.avatar.url if (profile and profile.avatar) else None,
            "presence": presence,
            "presence_label": presence_label,
            "status_msg": status_msg,
            "is_punched_in": is_punched_in,
            "punch_in_str": punch_in_str,
            "punch_out_str": punch_out_str,
            "team_key": team_key,
        }

        # Apply search query filter if provided
        if search_query:
            query_match = (
                search_query in member_data["full_name"].lower() or
                search_query in member_data["email"].lower() or
                search_query in member_data["designation"].lower() or
                search_query in member_data["department"].lower() or
                search_query in member_data["username"].lower()
            )
            if not query_match:
                continue

        teams_map[team_key]["members"].append(member_data)

    # Filter teams list for display
    display_teams = []
    for t in teams_config:
        t["member_count"] = len(t["members"])
        t["in_office_count"] = sum(1 for m in t["members"] if m["presence"] == "in_office")
        t["remote_count"] = sum(1 for m in t["members"] if m["presence"] in ["remote", "meeting"])
        t["leave_count"] = sum(1 for m in t["members"] if m["presence"] in ["on_leave", "out_of_office"])

        if selected_team == "all" or selected_team == t["key"]:
            display_teams.append(t)

    total_members = all_users.count()

    context = {
        "active_page": "teams",
        "today_attendance": today_attendance,
        "teams_config": teams_config,
        "display_teams": display_teams,
        "selected_team": selected_team,
        "search_query": search_query,
        "total_teams_count": len(teams_config),
        "total_members": total_members,
        "total_in_office": total_in_office,
        "total_remote": total_remote,
        "total_on_leave": total_on_leave,
    }
    return render(request, "accounts/teams.html", context)


@login_required(login_url="login")
def employees_export_excel_view(request):
    queryset, search_query, role_filter, status_filter, active_filter = _get_filtered_employees(request)
    today = timezone.localdate()
    today_attendances = {
        att.user_id: att for att in Attendance.objects.filter(date=today)
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employees Directory"
    ws.views.sheetView[0].showGridLines = True

    # Title styling
    title_font = Font(name="Arial", size=15, bold=True, color="7C5DFA")
    meta_font = Font(name="Arial", size=10, italic=True, color="64748B")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10, color="1E293B")
    bold_data_font = Font(name="Arial", size=10, bold=True, color="1E293B")

    header_fill = PatternFill(start_color="7C5DFA", end_color="7C5DFA", fill_type="solid")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Header rows
    ws.merge_cells("A1:J1")
    ws["A1"] = "ENTERPRISE HRMS PORTAL - EMPLOYEE DIRECTORY REPORT"
    ws["A1"].font = title_font
    ws["A1"].alignment = align_left

    ws.merge_cells("A2:J2")
    ws["A2"] = f"Generated on {timezone.now().strftime('%b %d, %Y at %I:%M %p')} | Total Records: {queryset.count()}"
    ws["A2"].font = meta_font
    ws["A2"].alignment = align_left

    headers = [
        "EMP ID",
        "Full Name",
        "Username",
        "Email Address",
        "System Role",
        "Live Presence",
        "Check-In",
        "Check-Out",
        "Joined Date",
        "Status",
    ]

    ws.append([])  # blank row 3
    ws.append(headers)  # row 4

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border

    # Data Rows
    for r_idx, emp in enumerate(queryset, start=5):
        fill = alt_fill if r_idx % 2 == 0 else white_fill
        
        emp_id = f"EMP-{emp.id:04d}"
        full_name = emp.get_full_name() or emp.username
        role_label, _, _ = _get_user_role_details(emp)
        
        presence_label = "In Office"
        if hasattr(emp, "work_status") and emp.work_status:
            presence_label = emp.work_status.get_status_display()

        att = today_attendances.get(emp.id)
        check_in_str = att.punch_in.strftime("%I:%M %p") if (att and att.punch_in) else ("Leave" if (att and att.status == "leave") else "-")
        check_out_str = att.punch_out.strftime("%I:%M %p") if (att and att.punch_out) else ("Active" if (att and att.is_punched_in) else "-")
        joined_str = emp.date_joined.strftime("%b %d, %Y") if emp.date_joined else "-"
        status_str = "Active" if emp.is_active else "Inactive"

        row_data = [
            emp_id,
            full_name,
            emp.username,
            emp.email or "-",
            role_label,
            presence_label,
            check_in_str,
            check_out_str,
            joined_str,
            status_str,
        ]

        ws.append(row_data)

        for c_idx in range(1, len(row_data) + 1):
            c = ws.cell(row=r_idx, column=c_idx)
            c.font = bold_data_font if c_idx == 1 else data_font
            c.fill = fill
            c.border = thin_border
            c.alignment = align_center if c_idx in [1, 5, 6, 7, 8, 9, 10] else align_left

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"employees_directory_{timezone.localdate().strftime('%Y%m%d')}.xlsx"
    response = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


@login_required(login_url="login")
def employees_export_pdf_view(request):
    queryset, search_query, role_filter, status_filter, active_filter = _get_filtered_employees(request)
    today = timezone.localdate()
    today_attendances = {
        att.user_id: att for att in Attendance.objects.filter(date=today)
    }

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=17,
        textColor=colors.HexColor("#1e293b"),
        spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "ReportSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=12,
    )
    table_cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        textColor=colors.HexColor("#1e293b"),
    )
    table_cell_bold = ParagraphStyle(
        "TableCellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=colors.HexColor("#1e293b"),
    )
    header_cell_style = ParagraphStyle(
        "HeaderCell",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        textColor=colors.white,
    )

    elements = [
        Paragraph("Enterprise HRMS Portal — Employee Directory", title_style),
        Paragraph(
            f"Report Generated: {timezone.now().strftime('%B %d, %Y at %I:%M %p')} | Total Staff Records: {queryset.count()}",
            subtitle_style,
        ),
        Spacer(1, 4),
    ]

    data = [
        [
            Paragraph("EMP ID", header_cell_style),
            Paragraph("Full Name", header_cell_style),
            Paragraph("Email Address", header_cell_style),
            Paragraph("System Role", header_cell_style),
            Paragraph("Presence", header_cell_style),
            Paragraph("Check-In", header_cell_style),
            Paragraph("Check-Out", header_cell_style),
            Paragraph("Joined Date", header_cell_style),
        ]
    ]

    for emp in queryset:
        emp_id = f"EMP-{emp.id:04d}"
        full_name = emp.get_full_name() or emp.username
        role_label, _, _ = _get_user_role_details(emp)
        presence_label = emp.work_status.get_status_display() if hasattr(emp, "work_status") and emp.work_status else "In Office"
        att = today_attendances.get(emp.id)
        check_in_str = att.punch_in.strftime("%I:%M %p") if (att and att.punch_in) else ("On Leave" if (att and att.status == "leave") else "-")
        check_out_str = att.punch_out.strftime("%I:%M %p") if (att and att.punch_out) else ("Active" if (att and att.is_punched_in) else "-")
        joined_str = emp.date_joined.strftime("%b %d, %Y") if emp.date_joined else "-"

        data.append([
            Paragraph(emp_id, table_cell_bold),
            Paragraph(full_name, table_cell_bold),
            Paragraph(emp.email or "-", table_cell_style),
            Paragraph(role_label, table_cell_style),
            Paragraph(presence_label, table_cell_style),
            Paragraph(check_in_str, table_cell_style),
            Paragraph(check_out_str, table_cell_style),
            Paragraph(joined_str, table_cell_style),
        ])

    table = Table(data, colWidths=[65, 120, 150, 90, 85, 75, 75, 75])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7c5dfa")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 12))
    elements.append(
        Paragraph(
            "Confidential Internal Document &bull; Generated by Enterprise HRMS Portal &bull; All Rights Reserved",
            ParagraphStyle("FooterNote", parent=styles["Normal"], fontSize=8, textColor=colors.HexColor("#94a3b8"), alignment=1),
        )
    )

    doc.build(elements)
    buf.seek(0)

    filename = f"employees_directory_{timezone.localdate().strftime('%Y%m%d')}.pdf"
    response = HttpResponse(buf.getvalue(), content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _ensure_employee_profile(user):
    profile, created = EmployeeProfile.objects.get_or_create(
        user=user,
        defaults={
            "phone": f"+91 98450 {user.id:05d}",
            "personal_email": f"{user.username}.personal@gmail.com",
            "department": "Engineering & Technology" if not user.is_staff else "Human Resources & Operations",
            "designation": "Staff Specialist" if not user.is_superuser else "Executive Lead",
            "work_location": "Headquarters (Bangalore)",
            "account_holder_name": user.get_full_name() or user.username,
        }
    )
    return profile


def _render_employee_profile(request, target_user):
    profile = _ensure_employee_profile(target_user)
    today = timezone.localdate()
    today_attendance = Attendance.objects.filter(user=target_user, date=today).first()
    
    # Presence Status
    work_status = getattr(target_user, "work_status", None)
    
    # Attendance summary
    present_days_count = Attendance.objects.filter(user=target_user, status="present").count()
    approved_leaves_count = LeaveRequest.objects.filter(user=target_user, status="approved").count()
    pending_leaves_count = LeaveRequest.objects.filter(user=target_user, status="pending").count()
    total_tasks_count = DailyTask.objects.filter(user=target_user).count()

    role_label, role_badge, _ = _get_user_role_details(target_user)

    is_own_profile = (request.user.id == target_user.id)

    context = {
        "active_page": "profile",
        "target_user": target_user,
        "profile": profile,
        "role_label": role_label,
        "role_badge": role_badge,
        "work_status": work_status,
        "today_attendance": today_attendance,
        "present_days_count": present_days_count,
        "approved_leaves_count": approved_leaves_count,
        "pending_leaves_count": pending_leaves_count,
        "total_tasks_count": total_tasks_count,
        "is_own_profile": is_own_profile,
    }
    return render(request, "accounts/profile.html", context)


@login_required(login_url="login")
def user_profile_view(request):
    """Current logged-in user's profile with all 7 information tabs."""
    return _render_employee_profile(request, target_user=request.user)


@login_required(login_url="login")
def employee_detail_profile_view(request, user_id):
    """View any employee profile by user_id."""
    target_user = get_object_or_404(UserModel, id=user_id)
    return _render_employee_profile(request, target_user=target_user)


@login_required(login_url="login")
def profile_update_view(request):
    if request.method == "POST":
        user_id = request.POST.get("user_id", request.user.id)
        try:
            target_user_id = int(user_id)
        except (ValueError, TypeError):
            target_user_id = request.user.id

        if not request.user.is_superuser and target_user_id != request.user.id:
            messages.error(request, "Unauthorized to modify other employee profiles.")
            return redirect("user_profile")

        target_user = get_object_or_404(UserModel, id=target_user_id)
        profile = _ensure_employee_profile(target_user)

        # Update User fields
        first_name = request.POST.get("first_name", "").strip()
        last_name = request.POST.get("last_name", "").strip()
        email = request.POST.get("email", "").strip()

        if first_name:
            target_user.first_name = first_name
        if last_name:
            target_user.last_name = last_name
        if email:
            target_user.email = email
        target_user.save()

        # Update Profile fields
        profile.phone = request.POST.get("phone", profile.phone).strip()
        profile.personal_email = request.POST.get("personal_email", profile.personal_email).strip()
        profile.gender = request.POST.get("gender", profile.gender).strip()
        profile.marital_status = request.POST.get("marital_status", profile.marital_status).strip()
        profile.blood_group = request.POST.get("blood_group", profile.blood_group).strip()
        profile.emergency_contact_name = request.POST.get("emergency_contact_name", profile.emergency_contact_name).strip()
        profile.emergency_contact_phone = request.POST.get("emergency_contact_phone", profile.emergency_contact_phone).strip()
        profile.emergency_relation = request.POST.get("emergency_relation", profile.emergency_relation).strip()
        profile.current_address = request.POST.get("current_address", profile.current_address).strip()
        profile.permanent_address = request.POST.get("permanent_address", profile.permanent_address).strip()

        # Admin editable fields
        if request.user.is_superuser or request.user.is_staff:
            profile.department = request.POST.get("department", profile.department).strip()
            profile.designation = request.POST.get("designation", profile.designation).strip()
            profile.work_location = request.POST.get("work_location", profile.work_location).strip()
            profile.reporting_manager = request.POST.get("reporting_manager", profile.reporting_manager).strip()
            profile.work_shift = request.POST.get("work_shift", profile.work_shift).strip()
            profile.bank_name = request.POST.get("bank_name", profile.bank_name).strip()
            profile.account_number = request.POST.get("account_number", profile.account_number).strip()
            profile.ifsc_code = request.POST.get("ifsc_code", profile.ifsc_code).strip()
            profile.pan_number = request.POST.get("pan_number", profile.pan_number).strip()

        profile.save()
        messages.success(request, f"Profile details for {target_user.get_full_name() or target_user.username} updated successfully!")

        if target_user.id == request.user.id:
            return redirect("user_profile")
        return redirect("employee_profile", user_id=target_user.id)

    return redirect("user_profile")


@login_required(login_url="login")
def profile_avatar_upload_view(request):
    """Handles instant avatar/profile picture upload from device file picker."""
    if request.method == "POST":
        user_id = request.POST.get("user_id", request.user.id)
        try:
            target_user_id = int(user_id)
        except (ValueError, TypeError):
            target_user_id = request.user.id

        if not request.user.is_superuser and target_user_id != request.user.id:
            return JsonResponse({"status": "error", "message": "Unauthorized to update this photo."}, status=403)

        target_user = get_object_or_404(UserModel, id=target_user_id)
        profile = _ensure_employee_profile(target_user)

        avatar_file = request.FILES.get("avatar")
        if avatar_file:
            profile.avatar = avatar_file
            profile.save()
            return JsonResponse({
                "status": "success",
                "message": "Profile picture updated successfully!",
                "avatar_url": profile.avatar.url,
            })
        else:
            return JsonResponse({"status": "error", "message": "No image file provided."}, status=400)

    return JsonResponse({"status": "error", "message": "Invalid request method."}, status=405)


@login_required(login_url="login")
def attendance_export_monthly_excel_view(request):
    """Exports comprehensive monthly attendance report for all employees as an Excel workbook (.xlsx)."""
    today = timezone.localdate()

    # Parse month and year from GET parameters (default to current month/year)
    try:
        year = int(request.GET.get("year", today.year))
    except (ValueError, TypeError):
        year = today.year

    try:
        month = int(request.GET.get("month", today.month))
    except (ValueError, TypeError):
        month = today.month

    if month < 1 or month > 12:
        month = today.month
    if year < 2000 or year > 2100:
        year = today.year

    month_name = calendar.month_name[month]
    num_days = calendar.monthrange(year, month)[1]
    start_date = date(year, month, 1)
    end_date = date(year, month, num_days)

    # Query all active employees
    users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("id")
    department_filter = request.GET.get("department", "").strip()
    if department_filter:
        users = users.filter(profile__department__iexact=department_filter)

    # Query attendances and approved leaves in this month
    attendances = Attendance.objects.filter(date__gte=start_date, date__lte=end_date).select_related("user")
    att_map = {(att.user_id, att.date): att for att in attendances}

    leaves = LeaveRequest.objects.filter(start_date__lte=end_date, end_date__gte=start_date, status="approved")
    leave_set = set()
    for l in leaves:
        l_curr = max(l.start_date, start_date)
        l_end = min(l.end_date, end_date)
        while l_curr <= l_end:
            leave_set.add((l.user_id, l_curr))
            l_curr += timedelta(days=1)

    # Create Workbook
    wb = openpyxl.Workbook()

    # Typography & Styles
    title_font = Font(name="Calibri", size=15, bold=True, color="1E293B")
    subtitle_font = Font(name="Calibri", size=10, bold=False, color="64748B")
    header_fill = PatternFill(start_color="4318FF", end_color="4318FF", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    regular_font = Font(name="Calibri", size=11, color="1E293B")

    # Status fills for matrix
    present_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    present_font = Font(name="Calibri", size=10, bold=True, color="166534")
    absent_fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
    absent_font = Font(name="Calibri", size=10, bold=True, color="9F1239")
    leave_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    leave_font = Font(name="Calibri", size=10, bold=True, color="92400E")
    weekend_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    weekend_font = Font(name="Calibri", size=10, bold=False, color="94A3B8")

    thin_border_side = Side(border_style="thin", color="E2E8F0")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom_side = Side(border_style="medium", color="4318FF")
    header_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

    # -------------------------------------------------------------
    # SHEET 1: Monthly Summary
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Monthly Summary"
    ws1.views.sheetView[0].showGridLines = True

    # Title Block
    ws1.merge_cells("A1:K1")
    ws1["A1"] = f"ENTERPRISE HRMS - MONTHLY ATTENDANCE REPORT ({month_name.upper()} {year})"
    ws1["A1"].font = title_font
    ws1["A1"].alignment = Alignment(vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:K2")
    ws1["A2"] = f"Period: {start_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')} | Generated: {timezone.localtime().strftime('%b %d, %Y %I:%M %p')} | Total Staff: {users.count()} Employees"
    ws1["A2"].font = subtitle_font
    ws1["A2"].alignment = Alignment(vertical="center")
    ws1.row_dimensions[2].height = 18

    ws1.append([])

    # Summary Headers (Row 4)
    summary_headers = [
        "EMP ID",
        "Employee Name",
        "Username",
        "Department",
        "Designation",
        "Total Days",
        "Present Days",
        "Absent Days",
        "Approved Leaves",
        "Total Hours",
        "Attendance %"
    ]
    ws1.append(summary_headers)
    ws1.row_dimensions[4].height = 26

    for col_idx, h in enumerate(summary_headers, 1):
        c = ws1.cell(row=4, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 6, 7, 8, 9, 10, 11] else "left", vertical="center")
        c.border = header_border

    tot_present_sum = 0
    tot_absent_sum = 0
    tot_leave_sum = 0
    tot_hours_sum = 0.0

    max_day = num_days if (year < today.year or (year == today.year and month < today.month)) else (today.day if (year == today.year and month == today.month) else num_days)

    for u_idx, u in enumerate(users, start=5):
        emp_id_str = f"EMP-{u.id:04d}"
        u_name = u.get_full_name() or u.username
        dept = u.profile.department if hasattr(u, "profile") and u.profile.department else "Engineering"
        desig = u.profile.designation if hasattr(u, "profile") and u.profile.designation else "Staff"

        p_count = 0
        a_count = 0
        l_count = 0
        total_mins = 0

        for d in range(1, max_day + 1):
            c_date = date(year, month, d)
            att = att_map.get((u.id, c_date))
            if att and att.punch_in:
                p_count += 1
                total_mins += att.duration_minutes
            elif (u.id, c_date) in leave_set:
                l_count += 1
            else:
                if c_date.weekday() < 5:
                    a_count += 1

        total_working_days = max(1, p_count + a_count + l_count)
        att_pct = round((p_count / total_working_days) * 100)
        hours_str = f"{total_mins // 60}h {total_mins % 60}m"
        hours_dec = round(total_mins / 60, 1)

        tot_present_sum += p_count
        tot_absent_sum += a_count
        tot_leave_sum += l_count
        tot_hours_sum += hours_dec

        row_data = [
            emp_id_str,
            u_name,
            f"@{u.username}",
            dept,
            desig,
            total_working_days,
            p_count,
            a_count,
            l_count,
            hours_str,
            f"{att_pct}%"
        ]
        ws1.append(row_data)
        ws1.row_dimensions[u_idx].height = 20

        is_even = (u_idx % 2 == 0)
        for col_idx in range(1, len(row_data) + 1):
            c = ws1.cell(row=u_idx, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if is_even:
                c.fill = zebra_fill
            if col_idx in [1, 6, 7, 8, 9, 10, 11]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    # Summary Total Row
    tot_row_idx = len(users) + 5
    avg_pct = round((tot_present_sum / max(1, tot_present_sum + tot_absent_sum + tot_leave_sum)) * 100) if (tot_present_sum + tot_absent_sum + tot_leave_sum) > 0 else 100
    total_row = [
        "TOTAL / AVG",
        f"{len(users)} Employees",
        "-",
        "-",
        "-",
        "-",
        tot_present_sum,
        tot_absent_sum,
        tot_leave_sum,
        f"{int(tot_hours_sum)}h",
        f"{avg_pct}%"
    ]
    ws1.append(total_row)
    ws1.row_dimensions[tot_row_idx].height = 22
    for col_idx in range(1, len(total_row) + 1):
        c = ws1.cell(row=tot_row_idx, column=col_idx)
        c.fill = total_fill
        c.font = total_font
        c.border = thin_border
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 6, 7, 8, 9, 10, 11] else "left", vertical="center")

    ws1.column_dimensions["A"].width = 13
    ws1.column_dimensions["B"].width = 24
    ws1.column_dimensions["C"].width = 18
    ws1.column_dimensions["D"].width = 24
    ws1.column_dimensions["E"].width = 24
    ws1.column_dimensions["F"].width = 13
    ws1.column_dimensions["G"].width = 14
    ws1.column_dimensions["H"].width = 14
    ws1.column_dimensions["I"].width = 16
    ws1.column_dimensions["J"].width = 14
    ws1.column_dimensions["K"].width = 15

    # -------------------------------------------------------------
    # SHEET 2: Daily Attendance Matrix (1..31 days)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Daily Matrix")
    ws2.views.sheetView[0].showGridLines = True

    # Title Block
    ws2.merge_cells(f"A1:{get_column_letter(num_days + 6)}1")
    ws2["A1"] = f"DAILY ATTENDANCE MATRIX - {month_name.upper()} {year}"
    ws2["A1"].font = title_font
    ws2["A1"].alignment = Alignment(vertical="center")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells(f"A2:{get_column_letter(num_days + 6)}2")
    ws2["A2"] = "Legend: P = Present | A = Absent | L = Approved Leave | WO = Weekend Off"
    ws2["A2"].font = subtitle_font
    ws2["A2"].alignment = Alignment(vertical="center")
    ws2.row_dimensions[2].height = 18

    ws2.append([])

    # Matrix Headers (Row 4)
    matrix_headers = ["EMP ID", "Employee Name", "Department"]
    for d in range(1, num_days + 1):
        c_date = date(year, month, d)
        day_abbr = c_date.strftime("%a")[:1]
        matrix_headers.append(f"{d:02d}\n{day_abbr}")
    matrix_headers.extend(["Present", "Absent", "Leave"])

    ws2.append(matrix_headers)
    ws2.row_dimensions[4].height = 28

    for col_idx, h in enumerate(matrix_headers, 1):
        c = ws2.cell(row=4, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border

    # Populate Matrix Rows
    for u_idx, u in enumerate(users, start=5):
        emp_id_str = f"EMP-{u.id:04d}"
        u_name = u.get_full_name() or u.username
        dept = u.profile.department if hasattr(u, "profile") and u.profile.department else "Engineering"

        m_p = 0
        m_a = 0
        m_l = 0

        row_vals = [emp_id_str, u_name, dept]

        for d in range(1, num_days + 1):
            c_date = date(year, month, d)
            is_weekend = (c_date.weekday() >= 5)
            att = att_map.get((u.id, c_date))

            if att and att.punch_in:
                val = "P"
                m_p += 1
            elif (u.id, c_date) in leave_set:
                val = "L"
                m_l += 1
            elif c_date <= today:
                if is_weekend:
                    val = "WO"
                else:
                    val = "A"
                    m_a += 1
            else:
                val = "-"

            row_vals.append(val)

        row_vals.extend([m_p, m_a, m_l])
        ws2.append(row_vals)
        ws2.row_dimensions[u_idx].height = 20

        # Apply cell styling for day columns
        for col_idx in range(1, len(row_vals) + 1):
            c = ws2.cell(row=u_idx, column=col_idx)
            c.border = thin_border
            if col_idx in [1, 2, 3]:
                c.font = regular_font
                c.alignment = Alignment(horizontal="left" if col_idx != 1 else "center", vertical="center")
            elif col_idx > num_days + 3:
                c.font = total_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = total_fill
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")
                day_val = c.value
                if day_val == "P":
                    c.fill = present_fill
                    c.font = present_font
                elif day_val == "A":
                    c.fill = absent_fill
                    c.font = absent_font
                elif day_val == "L":
                    c.fill = leave_fill
                    c.font = leave_font
                elif day_val == "WO":
                    c.fill = weekend_fill
                    c.font = weekend_font
                else:
                    c.font = Font(color="94A3B8")

    ws2.column_dimensions["A"].width = 13
    ws2.column_dimensions["B"].width = 22
    ws2.column_dimensions["C"].width = 22
    for d in range(1, num_days + 1):
        col_letter = get_column_letter(d + 3)
        ws2.column_dimensions[col_letter].width = 5.5
    ws2.column_dimensions[get_column_letter(num_days + 4)].width = 9
    ws2.column_dimensions[get_column_letter(num_days + 5)].width = 9
    ws2.column_dimensions[get_column_letter(num_days + 6)].width = 9

    # Generate HTTP response
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    filename = f"Monthly_Attendance_{month_name}_{year}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


@login_required(login_url="login")
def attendance_export_daily_excel_view(request):
    """Exports daily team attendance ledger for a given date as an Excel workbook (.xlsx)."""
    today = timezone.localdate()
    daily_date_str = request.GET.get("daily_date", "").strip()
    try:
        daily_date = datetime.strptime(daily_date_str, "%Y-%m-%d").date() if daily_date_str else today
    except ValueError:
        daily_date = today

    all_users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("first_name", "username")
    existing_daily_atts = {att.user_id: att for att in Attendance.objects.filter(date=daily_date).select_related("user", "user__profile")}
    approved_leaves = {lr.user_id: lr for lr in LeaveRequest.objects.filter(start_date__lte=daily_date, end_date__gte=daily_date, status="approved")}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Daily {daily_date.strftime('%b %d')}"
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:H1")
    ws["A1"] = f"DAILY TEAM ATTENDANCE LEDGER - {daily_date.strftime('%A, %B %d, %Y').upper()}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Generated On: {timezone.localtime().strftime('%b %d, %Y %I:%M %p')} | Organization: Enterprise HRMS"
    ws["A2"].font = Font(name="Calibri", size=10, color="64748B")
    ws.row_dimensions[2].height = 18

    headers = [
        "EMP ID",
        "Employee Name",
        "Department",
        "Punch In",
        "Punch Out",
        "Shift Duration",
        "Daily Status",
        "Shift Notes"
    ]
    ws.append(headers)
    ws.row_dimensions[4].height = 24

    header_fill = PatternFill(start_color="4318FF", end_color="4318FF", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 4, 5, 6, 7] else "left", vertical="center")

    thin_border = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for idx, u in enumerate(all_users, start=5):
        att = existing_daily_atts.get(u.id)
        if att and att.punch_in:
            stat = "PRESENT"
            p_in = timezone.localtime(att.punch_in).strftime("%I:%M %p")
            p_out = timezone.localtime(att.punch_out).strftime("%I:%M %p") if att.punch_out else "Active Now"
            dur = att.formatted_duration
            notes = att.notes or "Biometric shift recorded."
        elif u.id in approved_leaves:
            stat = "LEAVE"
            p_in = "--:--"
            p_out = "--:--"
            dur = "0h 0m"
            notes = f"Approved {approved_leaves[u.id].get_leave_type_display()}"
        else:
            stat = "ABSENT"
            p_in = "--:--"
            p_out = "--:--"
            dur = "0h 0m"
            notes = "Unrecorded shift - Absent"

        dept = u.profile.department if hasattr(u, "profile") and u.profile.department else "Engineering"
        row = [f"EMP-{u.id:04d}", u.get_full_name() or u.username, dept, p_in, p_out, dur, stat, notes]
        ws.append(row)
        ws.row_dimensions[idx].height = 20

        is_even = (idx % 2 == 0)
        for col_idx in range(1, len(row) + 1):
            c = ws.cell(row=idx, column=col_idx)
            c.font = Font(name="Calibri", size=11)
            c.border = thin_border
            if is_even:
                c.fill = zebra_fill
            if col_idx in [1, 4, 5, 6, 7]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            if col_idx == 7:
                if stat == "PRESENT":
                    c.font = Font(name="Calibri", size=10, bold=True, color="166534")
                    c.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                elif stat == "ABSENT":
                    c.font = Font(name="Calibri", size=10, bold=True, color="9F1239")
                    c.fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
                elif stat == "LEAVE":
                    c.font = Font(name="Calibri", size=10, bold=True, color="92400E")
                    c.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 30

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    filename = f"Daily_Attendance_{daily_date.strftime('%Y_%m_%d')}.xlsx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


@login_required(login_url="login")
def employees_export_excel_view(request):
    """Exports employee directory roster to Excel (.xlsx)."""
    users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("id")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employee Directory"
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws.merge_cells("A1:G1")
    ws["A1"] = "ENTERPRISE HRMS - EMPLOYEE DIRECTORY ROSTER"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws.row_dimensions[1].height = 25

    headers = ["EMP ID", "Employee Name", "Username", "Email Address", "Department", "Designation", "Joined Date"]
    ws.append(headers)
    ws.row_dimensions[2].height = 24

    header_fill = PatternFill(start_color="4318FF", end_color="4318FF", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 7] else "left", vertical="center")

    thin_border = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for idx, u in enumerate(users, start=3):
        dept = u.profile.department if hasattr(u, "profile") and u.profile.department else "Engineering"
        desig = u.profile.designation if hasattr(u, "profile") and u.profile.designation else "Staff"
        row = [f"EMP-{u.id:04d}", u.get_full_name() or u.username, f"@{u.username}", u.email or "-", dept, desig, u.date_joined.strftime("%b %d, %Y")]
        ws.append(row)
        ws.row_dimensions[idx].height = 20
        is_even = (idx % 2 == 0)
        for col_idx in range(1, len(row) + 1):
            c = ws.cell(row=idx, column=col_idx)
            c.font = Font(name="Calibri", size=11)
            c.border = thin_border
            if is_even:
                c.fill = zebra_fill
            if col_idx in [1, 7]:
                c.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = 'attachment; filename="Employee_Directory.xlsx"'
    wb.save(response)
    return response


@login_required(login_url="login")
def employees_export_pdf_view(request):
    """Exports employee directory roster to PDF (.pdf)."""
    import io
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    users = UserModel.objects.filter(is_active=True).select_related("profile", "work_status").order_by("id")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30,
    )
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="TitleStyle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#1E293B"),
    )
    subtitle_style = ParagraphStyle(
        name="SubtitleStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#64748B"),
    )
    cell_style = ParagraphStyle(
        name="CellStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#1E293B"),
    )
    header_cell_style = ParagraphStyle(
        name="HeaderCellStyle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=colors.white,
    )

    elements.append(Paragraph("ENTERPRISE HRMS - EMPLOYEE DIRECTORY ROSTER", title_style))
    elements.append(Paragraph(f"Generated On: {timezone.localtime().strftime('%b %d, %Y %I:%M %p')} | Total Staff: {users.count()} Employees", subtitle_style))
    elements.append(Spacer(1, 12))

    table_data = [[
        Paragraph("EMP ID", header_cell_style),
        Paragraph("Employee Name", header_cell_style),
        Paragraph("Username", header_cell_style),
        Paragraph("Email Address", header_cell_style),
        Paragraph("Department", header_cell_style),
        Paragraph("Designation", header_cell_style),
        Paragraph("Joined Date", header_cell_style),
    ]]

    for u in users:
        dept = u.profile.department if hasattr(u, "profile") and u.profile.department else "Engineering"
        desig = u.profile.designation if hasattr(u, "profile") and u.profile.designation else "Staff"
        table_data.append([
            Paragraph(f"EMP-{u.id:04d}", cell_style),
            Paragraph(u.get_full_name() or u.username, cell_style),
            Paragraph(f"@{u.username}", cell_style),
            Paragraph(u.email or "-", cell_style),
            Paragraph(dept, cell_style),
            Paragraph(desig, cell_style),
            Paragraph(u.date_joined.strftime("%b %d, %Y"), cell_style),
        ])

    table = Table(table_data, colWidths=[65, 130, 95, 150, 120, 120, 75])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4318FF")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 7),
        ("TOPPADDING", (0, 0), (-1, 0), 7),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("TOPPADDING", (0, 1), (-1, -1), 5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F8FAFC")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ]))

    elements.append(table)
    doc.build(elements)

    response = HttpResponse(buf.getvalue(), content_type="application/pdf")
    response["Content-Disposition"] = 'attachment; filename="Employee_Directory.pdf"'
    return response


@login_required
def holidays_view(request):
    today = timezone.localdate()
    current_year = 2026

    # 24 Structured Holidays for 2026 matching Company (12), Public (10), Optional (2)
    holidays_list = [
        {"name": "New Year's Day", "date": date(2026, 1, 1), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Global celebration marking the start of the year."},
        {"name": "Makar Sankranti / Pongal", "date": date(2026, 1, 14), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Harvest festival celebrated across India."},
        {"name": "Republic Day", "date": date(2026, 1, 26), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Honoring the Constitution of India."},
        {"name": "Maha Shivaratri", "date": date(2026, 2, 16), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Hindu festival honouring Shiva."},
        {"name": "Holi (Festival of Colours)", "date": date(2026, 3, 4), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Spring festival of colours."},
        {"name": "Ugadi / Gudi Padwa", "date": date(2026, 3, 19), "category": "Optional Holiday", "type": "Optional Holiday", "color": "orange", "description": "Traditional New Year day for Deccan region."},
        {"name": "Eid al-Fitr (Ramzan)", "date": date(2026, 3, 21), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Islamic festival celebrating end of Ramadan."},
        {"name": "Good Friday", "date": date(2026, 4, 3), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Christian commemoration of crucifixion of Jesus."},
        {"name": "Ambedkar Jayanti", "date": date(2026, 4, 14), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Dr. B. R. Ambedkar birthday commemoration."},
        {"name": "Labour Day / May Day", "date": date(2026, 5, 1), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "International Workers' Day."},
        {"name": "Buddha Purnima", "date": date(2026, 5, 31), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Commemorating birth of Gautama Buddha."},
        {"name": "Bakrid / Eid al-Adha", "date": date(2026, 5, 27), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Feast of the Sacrifice."},
        {"name": "Muharram", "date": date(2026, 6, 26), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Islamic New Year and remembrance."},
        {"name": "Independence Day", "date": date(2026, 8, 15), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "National Independence Day."},
        {"name": "Janmashtami", "date": date(2026, 8, 17), "category": "Optional Holiday", "type": "Optional Holiday", "color": "orange", "description": "Birth celebration of Lord Krishna."},
        {"name": "Onam (Thiruvonam)", "date": date(2026, 8, 29), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Harvest festival of Kerala."},
        {"name": "Teachers' Day", "date": date(2026, 9, 5), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "National observance celebrating educators."},
        {"name": "Ganesh Chaturthi", "date": date(2026, 9, 2), "category": "Company Holiday", "type": "Company Holiday", "color": "red", "description": "Festival celebrating arrival of Ganesha."},
        {"name": "Gandhi Jayanti", "date": date(2026, 10, 2), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Mahatma Gandhi birthday."},
        {"name": "Maha Navami / Dussehra", "date": date(2026, 10, 20), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Vijayadashami celebration of victory."},
        {"name": "Diwali (Deepavali)", "date": date(2026, 11, 8), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Festival of Lights."},
        {"name": "Govardhan Puja", "date": date(2026, 11, 9), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Post-Diwali celebration."},
        {"name": "Guru Nanak Jayanti", "date": date(2026, 11, 24), "category": "Public Holiday", "type": "Public Holiday", "color": "green", "description": "Birth anniversary of Guru Nanak Dev Ji."},
        {"name": "Christmas Day", "date": date(2026, 12, 25), "category": "Company Holiday", "type": "Company Holiday", "color": "purple", "description": "Celebration of the birth of Jesus Christ."},
    ]

    total_holidays = len(holidays_list)
    company_count = len([h for h in holidays_list if h["category"] == "Company Holiday"])
    public_count = len([h for h in holidays_list if h["category"] == "Public Holiday"])
    optional_count = len([h for h in holidays_list if h["category"] == "Optional Holiday"])

    upcoming_holidays = []
    past_holidays = []

    for h in holidays_list:
        h_date = h["date"]
        h["formatted_date"] = h_date.strftime("%A, %d %B %Y")
        h["day_name"] = h_date.strftime("%A")
        h["month_name"] = h_date.strftime("%B")
        h["month_abbr"] = h_date.strftime("%b").upper()
        h["day_num"] = h_date.strftime("%d")
        h["iso_date"] = h_date.strftime("%Y-%m-%d")
        days_diff = (h_date - today).days
        h["days_diff"] = days_diff

        if days_diff < 0:
            h["status"] = "past"
            h["badge_text"] = "Completed"
            past_holidays.append(h)
        elif days_diff == 0:
            h["status"] = "today"
            h["badge_text"] = "Today"
            upcoming_holidays.append(h)
        else:
            h["status"] = "upcoming"
            h["badge_text"] = f"In {days_diff} Days" if days_diff > 1 else "Tomorrow"
            upcoming_holidays.append(h)

    # Highlighted upcoming list (e.g. next 5-6 upcoming holidays starting from August 2026)
    highlight_upcoming = [
        {"month_abbr": "AUG", "day_num": "15", "name": "Independence Day", "full_date": "Saturday, 15 August 2026", "category": "Public Holiday", "color": "green"},
        {"month_abbr": "AUG", "day_num": "17", "name": "Janmashtami", "full_date": "Monday, 17 August 2026", "category": "Optional Holiday", "color": "orange"},
        {"month_abbr": "AUG", "day_num": "29", "name": "Onam", "full_date": "Saturday, 29 August 2026", "category": "Company Holiday", "color": "purple"},
        {"month_abbr": "SEP", "day_num": "05", "name": "Teachers' Day", "full_date": "Saturday, 05 September 2026", "category": "Public Holiday", "color": "green"},
        {"month_abbr": "SEP", "day_num": "02", "name": "Ganesh Chaturthi", "full_date": "Wednesday, 02 September 2026", "category": "Company Holiday", "color": "red"},
    ]

    import json
    holidays_json = json.dumps([
        {
            "name": h["name"],
            "date": h["iso_date"],
            "category": h["category"],
            "color": h["color"],
            "description": h["description"],
            "full_date": h["formatted_date"],
            "month_abbr": h["month_abbr"],
            "day_num": h["day_num"]
        }
        for h in holidays_list
    ])

    context = {
        "active_page": "holidays",
        "holidays": holidays_list,
        "holidays_json": holidays_json,
        "upcoming_holidays": upcoming_holidays,
        "highlight_upcoming": highlight_upcoming,
        "total_holidays": total_holidays,
        "company_count": company_count,
        "public_count": public_count,
        "optional_count": optional_count,
        "current_year": current_year,
    }
    return render(request, "accounts/holidays.html", context)


def format_inr(number):
    try:
        num = int(round(float(number)))
        s = str(num)
        if len(s) <= 3:
            return s
        last_three = s[-3:]
        remaining = s[:-3]
        parts = []
        while remaining:
            parts.append(remaining[-2:])
            remaining = remaining[:-2]
        parts.reverse()
        return ",".join(parts) + "," + last_three
    except Exception:
        return str(number)


@login_required
def payroll_view(request):
    profile, _ = EmployeeProfile.objects.get_or_create(user=request.user)
    today = timezone.localdate()
    all_profiles = list(EmployeeProfile.objects.select_related("user").all())

    # Aggregate department statistics directly from DB
    dept_map = {}
    total_db_employees = len(all_profiles)
    total_gross_db = sum(p.gross_monthly for p in all_profiles) if all_profiles else 0
    total_deductions_db = sum(p.total_deductions_monthly for p in all_profiles) if all_profiles else 0
    total_net_db = sum(p.net_monthly for p in all_profiles) if all_profiles else 0

    # Group profiles by department
    for p in all_profiles:
        raw_dept = p.department.strip() if p.department else "General Operations"
        if raw_dept.lower() in ["engineering", "engineering & product", "product & engineering"]:
            d_name = "Engineering & Product"
        elif raw_dept.lower() in ["design", "design & quality assurance", "qa & design"]:
            d_name = "Design & Quality Assurance"
        elif raw_dept.lower() in ["human resources", "hr", "human resources & operations"]:
            d_name = "Human Resources & Operations"
        elif raw_dept.lower() in ["executive management", "executive", "management"]:
            d_name = "Executive Management"
        else:
            d_name = raw_dept

        if d_name not in dept_map:
            dept_map[d_name] = {
                "id": f"dept-{len(dept_map)+1}",
                "name": d_name,
                "employees_count": 0,
                "gross_salary": 0,
                "deductions": 0,
                "net_payout": 0,
                "status": "Processed",
            }
        dept_map[d_name]["employees_count"] += 1
        dept_map[d_name]["gross_salary"] += p.gross_monthly
        dept_map[d_name]["deductions"] += p.total_deductions_monthly
        dept_map[d_name]["net_payout"] += p.net_monthly

    departments_data = []
    for d_name, data in sorted(dept_map.items()):
        departments_data.append({
            "id": data["id"],
            "name": data["name"],
            "employees_count": data["employees_count"],
            "gross_formatted": format_inr(data["gross_salary"]),
            "deductions_formatted": format_inr(data["deductions"]),
            "net_formatted": format_inr(data["net_payout"]),
            "status": data["status"],
        })

    # DB Itemized payslips for all active users
    month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    curr_month_name = month_names[today.month - 1]
    curr_year = today.year

    db_payslips = []
    for p in all_profiles:
        db_payslips.append({
            "slip_id": f"PAY-{curr_year}{today.month:02d}-{p.user.id:04d}",
            "employee_name": p.user.get_full_name() or p.user.username,
            "department": p.department,
            "designation": p.designation,
            "period": f"{curr_month_name} {curr_year}",
            "gross_formatted": format_inr(p.gross_monthly),
            "deductions_formatted": format_inr(p.total_deductions_monthly),
            "net_formatted": format_inr(p.net_monthly),
            "status": "Paid",
        })

    # Deductions overview computed from DB
    pf_total = sum(p.pf_deduction_monthly for p in all_profiles) if all_profiles else 0
    tax_total = sum(p.tax_deduction_monthly for p in all_profiles) if all_profiles else 0

    deductions_summary = [
        {"component": "Provident Fund (PF)", "type": "Statutory Deduction", "rate": "12% of Basic", "total_formatted": format_inr(pf_total), "compliance": "EPFO 1952"},
        {"component": "Tax Deducted at Source (TDS)", "type": "Income Tax", "rate": "Slab Based", "total_formatted": format_inr(tax_total), "compliance": "Income Tax Act"},
        {"component": "Professional Tax (PT)", "type": "State Tax", "rate": "INR 200 / Month", "total_formatted": format_inr(total_db_employees * 200), "compliance": "State Govt"},
        {"component": "Health Insurance (ESI / Mediclaim)", "type": "Voluntary / Group", "rate": "Fixed", "total_formatted": format_inr(total_db_employees * 500), "compliance": "ESIC"},
    ]

    # Salary components overview
    salary_components = [
        {"name": "Basic Salary", "type": "Earnings", "percentage": "50% of CTC", "taxable": "Fully Taxable", "frequency": "Monthly"},
        {"name": "House Rent Allowance (HRA)", "type": "Earnings", "percentage": "25% of CTC", "taxable": "Exemption u/s 10(13A)", "frequency": "Monthly"},
        {"name": "Special Allowance", "type": "Earnings", "percentage": "15% of CTC", "taxable": "Fully Taxable", "frequency": "Monthly"},
        {"name": "Conveyance Allowance", "type": "Earnings", "percentage": "5% of CTC", "taxable": "Taxable", "frequency": "Monthly"},
        {"name": "Performance Bonus / Incentive", "type": "Variable Earnings", "percentage": "5% of CTC", "taxable": "Fully Taxable", "frequency": "Quarterly / Annual"},
    ]

    # Historical runs
    gross_flt = float(total_gross_db)
    ded_flt = float(total_deductions_db)
    net_flt = float(total_net_db)

    history_runs = [
        {"period": f"July {curr_year}", "disbursed_on": f"28 Jul {curr_year}", "employees": total_db_employees, "gross": format_inr(gross_flt * 0.98), "deductions": format_inr(ded_flt * 0.98), "net": format_inr(net_flt * 0.98), "status": "Completed"},
        {"period": f"June {curr_year}", "disbursed_on": f"28 Jun {curr_year}", "employees": total_db_employees, "gross": format_inr(gross_flt * 0.96), "deductions": format_inr(ded_flt * 0.96), "net": format_inr(net_flt * 0.96), "status": "Completed"},
        {"period": f"May {curr_year}", "disbursed_on": f"28 May {curr_year}", "employees": total_db_employees, "gross": format_inr(gross_flt * 0.95), "deductions": format_inr(ded_flt * 0.95), "net": format_inr(net_flt * 0.95), "status": "Completed"},
        {"period": f"April {curr_year}", "disbursed_on": f"28 Apr {curr_year}", "employees": total_db_employees, "gross": format_inr(gross_flt * 0.93), "deductions": format_inr(ded_flt * 0.93), "net": format_inr(net_flt * 0.93), "status": "Completed"},
    ]

    context = {
        "active_page": "payroll",
        "profile": profile,
        "total_employees": total_db_employees,
        "total_gross_formatted": format_inr(total_gross_db),
        "total_deductions_formatted": format_inr(total_deductions_db),
        "total_net_formatted": format_inr(total_net_db),
        "current_period_short": f"{curr_month_name[:3].upper()} {curr_year}",
        "current_period_full": f"{curr_month_name.upper()} {curr_year}",
        "departments": departments_data,
        "db_payslips": db_payslips,
        "deductions_summary": deductions_summary,
        "salary_components": salary_components,
        "history_runs": history_runs,
    }
    return render(request, "accounts/payroll.html", context)




