from django.test import TestCase, Client
from django.urls import reverse
from django.contrib.auth import get_user_model
from django.utils import timezone
from accounts.models import Attendance, DailyTask, LeaveRequest, PermissionRequest, UserStatus

UserModel = get_user_model()


class FullPortalAndTaskTests(TestCase):
    def setUp(self):
        self.client = Client()
        self.email = "sridharmukesh07@gmail.com"
        self.username = "sridhar"
        self.password = "SecurePass123!"
        self.user = UserModel.objects.create_user(
            username=self.username,
            email=self.email,
            password=self.password,
            first_name="Sridhar",
            last_name="Arumugam",
        )
        self.client.login(username=self.username, password=self.password)

    def test_dashboard_renders_with_modular_sidebar_and_settings(self):
        response = self.client.get(reverse("home"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "SRIDHAR")
        self.assertContains(response, "sridharmukesh07@gmail.com")
        self.assertContains(response, "APPLICATIONS")
        self.assertContains(response, "Leave's")
        self.assertContains(response, "Status Board")
        self.assertContains(response, "Daily Task")
        self.assertContains(response, "Settings")
        self.assertContains(response, "HOURS WORKED")
        self.assertContains(response, "7. UPCOMING")
        self.assertContains(response, "Tasks")
        self.assertContains(response, "Birthdays")
        self.assertContains(response, "Holidays")
        self.assertContains(response, "4. PERFORMANCE")
        self.assertContains(response, "Best Employee")
        self.assertContains(response, "Best Team")
        self.assertContains(response, "Overall")

    def test_task_management_view_and_filtering(self):
        DailyTask.objects.create(user=self.user, title="Pending Task", status="pending")
        DailyTask.objects.create(user=self.user, title="In Progress Task", status="in_progress")
        DailyTask.objects.create(user=self.user, title="Completed Task", status="completed")

        response = self.client.get(reverse("tasks"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Daily Task Efficiency")
        self.assertContains(response, "Pending Task")

        # Status filter
        filter_resp = self.client.get(reverse("tasks") + "?status=in_progress")
        self.assertEqual(filter_resp.status_code, 200)
        self.assertEqual(filter_resp.context["selected_status"], "in_progress")

    def test_quick_inline_task_creation(self):
        response = self.client.post(
            reverse("task_create"),
            {
                "quick_add": "1",
                "title": "Quick inline task test",
                "priority": "high",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        task = DailyTask.objects.filter(user=self.user, title="Quick inline task test").first()
        self.assertIsNotNone(task)
        self.assertEqual(task.priority, "high")

    def test_profile_update_view(self):
        response = self.client.post(
            reverse("profile_update"),
            {
                "first_name": "Sridhar Updated",
                "last_name": "Arumugam Updated",
                "email": "sridhar.new@example.com",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        self.user.refresh_from_db()
        self.assertEqual(self.user.first_name, "Sridhar Updated")
        self.assertEqual(self.user.email, "sridhar.new@example.com")

    def test_password_change_view(self):
        response = self.client.post(
            reverse("password_change"),
            {
                "old_password": "SecurePass123!",
                "new_password1": "NewSecurePass456!",
                "new_password2": "NewSecurePass456!",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password("NewSecurePass456!"))

    def test_settings_page_render_with_appearance(self):
        response = self.client.get(reverse("settings"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Appearance & Accent Theme")
        self.assertContains(response, "Personal Profile Details")
        self.assertContains(response, "Password & Security")

    def test_status_board_update(self):
        # Update
        response = self.client.post(
            reverse("status_board"),
            {
                "status": "remote",
                "status_message": "Working from Bangalore hub",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        status = UserStatus.objects.get(user=self.user)
        self.assertEqual(status.status, "remote")
        self.assertEqual(status.status_message, "Working from Bangalore hub")

        # Render & Filter
        get_resp = self.client.get(reverse("status_board"))
        self.assertEqual(get_resp.status_code, 200)
        self.assertContains(get_resp, "Team Presence Broadcast")
        self.assertContains(get_resp, "Working from Bangalore hub")

        filter_resp = self.client.get(reverse("status_board") + "?filter=remote")
        self.assertEqual(filter_resp.status_code, 200)
        self.assertEqual(filter_resp.context["selected_filter"], "remote")

        # Delete / Clear Status Note
        del_resp = self.client.post(
            reverse("status_board"),
            {
                "action": "delete",
                "target_id": self.user.id,
            },
            follow=True,
        )
        self.assertEqual(del_resp.status_code, 200)
        status.refresh_from_db()
        self.assertEqual(status.status_message, "")

    def test_leave_application(self):
        today = timezone.localdate()
        response = self.client.post(
            reverse("leave_apply"),
            {
                "leave_type": "casual",
                "start_date": today.strftime("%Y-%m-%d"),
                "end_date": (today + timezone.timedelta(days=2)).strftime("%Y-%m-%d"),
                "reason": "Personal work",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(LeaveRequest.objects.filter(user=self.user, leave_type="casual").exists())

    def test_permissions_page_view(self):
        response = self.client.get(reverse("permissions"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Permissions & Gate Pass Management")

    def test_attendance_page_filters_and_pagination(self):
        today = timezone.localdate()
        # Create records with 4 statuses
        Attendance.objects.create(user=self.user, date=today, status="present")
        Attendance.objects.create(user=self.user, date=today - timezone.timedelta(days=1), status="present", notes="Late check-in")
        Attendance.objects.create(user=self.user, date=today - timezone.timedelta(days=2), status="absent")
        Attendance.objects.create(user=self.user, date=today - timezone.timedelta(days=3), status="leave")
        Attendance.objects.create(user=self.user, date=today - timezone.timedelta(days=4), status="lop")

        # Test base page
        response = self.client.get(reverse("attendance"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Present")
        self.assertContains(response, "Absent")
        self.assertContains(response, "Leave")
        self.assertContains(response, "Loss of Pay")

        # Test filter by present
        response_present = self.client.get(reverse("attendance") + "?status=present")
        self.assertEqual(response_present.status_code, 200)
        self.assertEqual(response_present.context["selected_status"], "present")

        # Test filter by lop
        response_lop = self.client.get(reverse("attendance") + "?status=lop")
        self.assertEqual(response_lop.status_code, 200)
        self.assertEqual(response_lop.context["selected_status"], "lop")

        # Test filter by leave
        response_leave = self.client.get(reverse("attendance") + "?status=leave")
        self.assertEqual(response_leave.status_code, 200)
        self.assertEqual(response_leave.context["selected_status"], "leave")

        # Test filter by absent
        response_absent = self.client.get(reverse("attendance") + "?status=absent")
        self.assertEqual(response_absent.status_code, 200)
        self.assertEqual(response_absent.context["selected_status"], "absent")

    def test_task_edit_crud(self):
        task = DailyTask.objects.create(user=self.user, title="Original Task", status="pending")
        response = self.client.post(
            reverse("task_edit", args=[task.id]),
            {
                "title": "Updated Task Title",
                "priority": "high",
                "status": "in_progress",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        task.refresh_from_db()
        self.assertEqual(task.title, "Updated Task Title")
        self.assertEqual(task.priority, "high")
        self.assertEqual(task.status, "in_progress")

    def test_leave_edit_and_delete_crud(self):
        today = timezone.localdate()
        leave = LeaveRequest.objects.create(
            user=self.user,
            leave_type="casual",
            start_date=today,
            end_date=today + timezone.timedelta(days=1),
            reason="Original reason",
        )
        # Edit
        response = self.client.post(
            reverse("leave_edit", args=[leave.id]),
            {
                "leave_type": "sick",
                "start_date": today.strftime("%Y-%m-%d"),
                "end_date": (today + timezone.timedelta(days=2)).strftime("%Y-%m-%d"),
                "reason": "Updated medical reason",
            },
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        leave.refresh_from_db()
        self.assertEqual(leave.leave_type, "sick")
        self.assertEqual(leave.reason, "Updated medical reason")

        # Delete / Cancel
        del_resp = self.client.post(reverse("leave_delete", args=[leave.id]), follow=True)
        self.assertEqual(del_resp.status_code, 200)
        self.assertFalse(LeaveRequest.objects.filter(id=leave.id).exists())

    def test_employees_directory_view_and_filtering(self):
        response = self.client.get(reverse("employees_directory"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Employee Directory & Staff Roster")
        self.assertContains(response, "sridhar")

        # Search filter
        search_resp = self.client.get(reverse("employees_directory") + "?q=sridhar")
        self.assertEqual(search_resp.status_code, 200)
        self.assertContains(search_resp, "sridhar")

        # Role filter
        role_resp = self.client.get(reverse("employees_directory") + "?role=superadmin")
        self.assertEqual(role_resp.status_code, 200)

    def test_employees_directory_pagination(self):
        # Create additional users to test pagination
        for i in range(10):
            UserModel.objects.get_or_create(
                username=f"page_user_{i}",
                defaults={"email": f"page_{i}@company.com", "first_name": f"PageUser{i}"}
            )
        response = self.client.get(reverse("employees_directory") + "?page=2")
        self.assertEqual(response.status_code, 200)
        self.assertIn("page_obj", response.context)
        self.assertEqual(response.context["page_obj"].number, 2)

    def test_employees_export_excel(self):
        response = self.client.get(reverse("employees_export_excel"))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertTrue(len(response.content) > 1000)

    def test_employees_export_pdf(self):
        response = self.client.get(reverse("employees_export_pdf"))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/pdf")
        self.assertTrue(len(response.content) > 1000)

    def test_user_profile_view_renders_all_tabs(self):
        response = self.client.get(reverse("user_profile"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Sridhar Arumugam")
        self.assertContains(response, "panel-employee")
        self.assertContains(response, "panel-personal")
        self.assertContains(response, "panel-finance")
        self.assertContains(response, "panel-salary")
        self.assertContains(response, "panel-documents")
        self.assertContains(response, "panel-increments")
        self.assertContains(response, "panel-experience")

    def test_employee_detail_profile_view(self):
        other_user = UserModel.objects.create_user(
            username="alex.dev",
            email="alex.dev@company.com",
            first_name="Alex",
            last_name="Developer",
        )
        response = self.client.get(reverse("employee_profile", kwargs={"user_id": other_user.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Alex Developer")
        self.assertContains(response, "EMP-")

    def test_profile_update_view(self):
        resp = self.client.post(
            reverse("profile_update"),
            {
                "user_id": self.user.id,
                "first_name": "Sridhar",
                "last_name": "Updated",
                "email": "sridhar.updated@gmail.com",
                "phone": "+91 99999 88888",
                "personal_email": "sridhar.personal@gmail.com",
                "gender": "Male",
                "marital_status": "Married",
                "blood_group": "O+",
                "emergency_contact_name": "Priya",
                "emergency_contact_phone": "+91 88888 77777",
                "emergency_relation": "Spouse",
                "current_address": "Indiranagar, Bangalore",
                "permanent_address": "Indiranagar, Bangalore",
            },
            follow=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.user.refresh_from_db()
        self.assertEqual(self.user.last_name, "Updated")
        self.assertEqual(self.user.profile.phone, "+91 99999 88888")
        self.assertEqual(self.user.profile.current_address, "Indiranagar, Bangalore")

    def test_punch_in_and_punch_out(self):
        # 1. Test Punch In
        resp_in = self.client.post(
            reverse("punch_attendance"),
            {"action": "punch_in"},
            follow=True,
        )
        self.assertEqual(resp_in.status_code, 200)
        today = timezone.localdate()
        att = Attendance.objects.get(user=self.user, date=today)
        self.assertIsNotNone(att.punch_in)
        self.assertIsNone(att.punch_out)
        self.assertTrue(att.is_punched_in)

        # 2. Test Punch Out
        resp_out = self.client.post(
            reverse("punch_attendance"),
            {"action": "punch_out"},
            follow=True,
        )
        self.assertEqual(resp_out.status_code, 200)
        att.refresh_from_db()
        self.assertIsNotNone(att.punch_out)
        self.assertTrue(att.is_completed)
        self.assertGreaterEqual(att.duration.total_seconds(), 0)

    def test_profile_avatar_upload(self):
        from django.core.files.uploadedfile import SimpleUploadedFile
        test_image = SimpleUploadedFile(
            name="test_avatar.png",
            content=b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15c4\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82",
            content_type="image/png",
        )
        resp = self.client.post(
            reverse("profile_avatar_upload"),
            {"avatar": test_image, "user_id": self.user.id},
        )
        self.assertEqual(resp.status_code, 200)
        data = resp.json()
        self.assertEqual(data["status"], "success")
        self.user.refresh_from_db()
        self.assertTrue(bool(self.user.profile.avatar))

    def test_permissions_view_and_apply(self):
        # 1. Test view renders correctly
        resp = self.client.get(reverse("permissions"))
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "Permissions & Gate Pass Management")
        self.assertContains(resp, "Monthly Available Quota")

        # 2. Test apply for permission
        resp_apply = self.client.post(
            reverse("permission_apply"),
            {
                "permission_type": "late_entry",
                "date": timezone.localdate().strftime("%Y-%m-%d"),
                "duration_hours": "1.5",
                "reason": "Traffic delay on highway",
            },
            follow=True,
        )
        self.assertEqual(resp_apply.status_code, 200)
        perm = PermissionRequest.objects.filter(user=self.user, permission_type="late_entry").first()
        self.assertIsNotNone(perm)
        self.assertEqual(perm.status, "pending")
        self.assertEqual(float(perm.duration_hours), 1.5)

    def test_permission_edit_and_delete(self):
        perm = PermissionRequest.objects.create(
            user=self.user,
            permission_type="early_exit",
            date=timezone.localdate(),
            duration_hours=2.0,
            reason="Doctor visit",
            status="pending",
        )
        # Edit
        resp_edit = self.client.post(
            reverse("permission_edit", kwargs={"pk": perm.pk}),
            {
                "permission_type": "early_exit",
                "date": timezone.localdate().strftime("%Y-%m-%d"),
                "duration_hours": "1.0",
                "reason": "Doctor visit rescheduled earlier",
            },
            follow=True,
        )
        self.assertEqual(resp_edit.status_code, 200)
        perm.refresh_from_db()
        self.assertEqual(float(perm.duration_hours), 1.0)
        self.assertIn("rescheduled", perm.reason)

        # Delete
        resp_del = self.client.post(
            reverse("permission_delete", kwargs={"pk": perm.pk}),
            follow=True,
        )
        self.assertEqual(resp_del.status_code, 200)
        self.assertFalse(PermissionRequest.objects.filter(pk=perm.pk).exists())

    def test_permission_admin_status_update(self):
        self.user.is_superuser = True
        self.user.is_staff = True
        self.user.save()

        perm = PermissionRequest.objects.create(
            user=self.user,
            permission_type="on_duty",
            date=timezone.localdate(),
            duration_hours=3.0,
            reason="Client on-site deployment",
            status="pending",
        )
        # Approve
        resp = self.client.post(
            reverse("permission_status_update", kwargs={"pk": perm.pk}),
            {"status": "approved"},
            follow=True,
        )
        self.assertEqual(resp.status_code, 200)
        perm.refresh_from_db()
        self.assertEqual(perm.status, "approved")

    def test_teams_view_and_filtering(self):
        # 1. Base Teams page render
        response = self.client.get(reverse("teams"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Teams & Functional Units")
        self.assertContains(response, "Super Admin")
        self.assertContains(response, "Admin")
        self.assertContains(response, "HR Manager")
        self.assertContains(response, "HR Executive")
        self.assertContains(response, "Manager")
        self.assertContains(response, "Development")
        self.assertContains(response, "Finance")
        self.assertContains(response, "Support")

        # 2. Filter by specific team (e.g. development)
        dev_resp = self.client.get(reverse("teams") + "?team=development")
        self.assertEqual(dev_resp.status_code, 200)
        self.assertEqual(dev_resp.context["selected_team"], "development")

        # 3. Filter by search query
        search_resp = self.client.get(reverse("teams") + "?q=sridhar")
        self.assertEqual(search_resp.status_code, 200)
        self.assertEqual(search_resp.context["search_query"], "sridhar")






