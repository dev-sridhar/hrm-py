from django.contrib import admin
from .models import Attendance, DailyTask, LeaveRequest, UserStatus, PermissionRequest, EmployeeProfile


@admin.register(Attendance)
class AttendanceAdmin(admin.ModelAdmin):
    list_display = ("user", "date", "punch_in", "punch_out", "status", "formatted_duration")
    list_filter = ("status", "date")
    search_fields = ("user__username", "user__email", "notes")
    date_hierarchy = "date"


@admin.register(DailyTask)
class DailyTaskAdmin(admin.ModelAdmin):
    list_display = ("title", "user", "date", "due_date", "priority", "status")
    list_filter = ("status", "priority", "date")
    search_fields = ("title", "description", "user__username", "user__email")
    date_hierarchy = "date"


@admin.register(PermissionRequest)
class PermissionRequestAdmin(admin.ModelAdmin):
    list_display = ("user", "permission_type", "date", "duration_hours", "status", "applied_at")
    list_filter = ("status", "permission_type", "date")
    search_fields = ("user__username", "reason")


@admin.register(LeaveRequest)
class LeaveRequestAdmin(admin.ModelAdmin):
    list_display = ("user", "leave_type", "start_date", "end_date", "status", "applied_at")
    list_filter = ("status", "leave_type")
    search_fields = ("user__username", "reason")


@admin.register(UserStatus)
class UserStatusAdmin(admin.ModelAdmin):
    list_display = ("user", "status", "status_message", "updated_at")
    list_filter = ("status",)
    search_fields = ("user__username", "status_message")
