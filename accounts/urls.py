from django.urls import path
from . import views

urlpatterns = [
    # Auth & Dashboard
    path("", views.home_view, name="home"),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("register/", views.register_view, name="register"),

    # Settings, Profile & Password
    path("settings/", views.settings_view, name="settings"),
    path("profile/", views.user_profile_view, name="user_profile"),
    path("profile/update/", views.profile_update_view, name="profile_update"),
    path("profile/avatar/upload/", views.profile_avatar_upload_view, name="profile_avatar_upload"),
    path("employees/<int:user_id>/profile/", views.employee_detail_profile_view, name="employee_profile"),
    path("password/change/", views.password_change_view, name="password_change"),

    # Status Board, Employees Directory & Teams
    path("status-board/", views.status_board_view, name="status_board"),
    path("employees/", views.employees_directory_view, name="employees_directory"),
    path("teams/", views.teams_view, name="teams"),
    path("employees/export/excel/", views.employees_export_excel_view, name="employees_export_excel"),
    path("employees/export/pdf/", views.employees_export_pdf_view, name="employees_export_pdf"),

    # Attendance & Export Reports
    path("attendance/", views.attendance_view, name="attendance"),
    path("attendance/punch/", views.punch_attendance_view, name="punch_attendance"),
    path("attendance/export/monthly/", views.attendance_export_monthly_excel_view, name="attendance_export_monthly"),
    path("attendance/export/daily/", views.attendance_export_daily_excel_view, name="attendance_export_daily"),

    # Daily Tasks
    path("tasks/", views.tasks_view, name="tasks"),
    path("tasks/create/", views.task_create_view, name="task_create"),
    path("tasks/<int:pk>/edit/", views.task_edit_view, name="task_edit"),
    path("tasks/<int:pk>/status/", views.task_update_status_view, name="task_update_status"),
    path("tasks/<int:pk>/delete/", views.task_delete_view, name="task_delete"),
    path("tasks/export/excel/", views.tasks_export_excel_view, name="tasks_export_excel"),

    # Applications - Leaves, Permissions, Holidays & Payroll
    path("applications/leaves/", views.leaves_view, name="leaves"),
    path("applications/leaves/apply/", views.leave_apply_view, name="leave_apply"),
    path("applications/leaves/<int:pk>/status/", views.leave_status_update_view, name="leave_status_update"),
    path("applications/leaves/<int:pk>/edit/", views.leave_edit_view, name="leave_edit"),
    path("applications/leaves/<int:pk>/delete/", views.leave_delete_view, name="leave_delete"),
    path("applications/permissions/", views.permissions_view, name="permissions"),
    path("applications/permissions/apply/", views.permission_apply_view, name="permission_apply"),
    path("applications/permissions/<int:pk>/status/", views.permission_status_update_view, name="permission_status_update"),
    path("applications/permissions/<int:pk>/edit/", views.permission_edit_view, name="permission_edit"),
    path("applications/permissions/<int:pk>/delete/", views.permission_delete_view, name="permission_delete"),
    path("applications/holidays/", views.holidays_view, name="holidays"),
    path("applications/payroll/", views.payroll_view, name="payroll"),
]
