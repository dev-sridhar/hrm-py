from django import template

register = template.Library()


@register.inclusion_tag("accounts/components/sidebar.html", takes_context=True)
def render_sidebar(context, active_page=None):
    """
    Component Tag/Class for rendering the Left Modular Floating Sidebar.
    Encapsulates active navigation resolution and user context.
    """
    request = context.get("request")
    current_active = active_page or context.get("active_page", "")

    # Auto-detect active page from request resolver if not explicitly provided
    if not current_active and request and hasattr(request, "resolver_match") and request.resolver_match:
        url_name = request.resolver_match.url_name
        if url_name == "home":
            current_active = "dashboard"
        elif url_name in ["employees_directory", "employee_profile"]:
            current_active = "employees"
        elif url_name == "teams":
            current_active = "teams"
        elif url_name in ["attendance", "punch_attendance"]:
            current_active = "attendance"
        elif url_name in ["tasks", "task_create", "task_update_status", "task_delete"]:
            current_active = "tasks"
        elif url_name in ["leaves", "leave_apply", "leave_edit", "leave_delete"]:
            current_active = "leaves"
        elif url_name in ["permissions", "permission_request", "permission_edit", "permission_delete"]:
            current_active = "permissions"
        elif url_name == "holidays":
            current_active = "holidays"
        elif url_name == "payroll":
            current_active = "payroll"
        elif url_name == "status_board":
            current_active = "status_board"
        elif url_name in ["settings", "user_profile"]:
            current_active = "settings" if url_name == "settings" else "profile"

    return {
        "request": request,
        "user": getattr(request, "user", None),
        "active_page": current_active,
    }

