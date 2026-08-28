from django.db import models
from django.contrib.auth import get_user_model
from django.utils import timezone
from datetime import datetime, time, timedelta

UserModel = get_user_model()


class Attendance(models.Model):
    STATUS_CHOICES = [
        ("present", "Present"),
        ("absent", "Absent"),
        ("leave", "Leave"),
        ("lop", "LOP"),
    ]

    user = models.ForeignKey(
        UserModel,
        on_delete=models.CASCADE,
        related_name="attendances",
    )
    date = models.DateField(default=timezone.localdate)
    punch_in = models.DateTimeField(null=True, blank=True)
    punch_out = models.DateTimeField(null=True, blank=True)
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="present",
    )
    notes = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-date", "-created_at"]
        constraints = [
            models.UniqueConstraint(
                fields=["user", "date"],
                name="unique_user_daily_attendance",
            )
        ]

    def __str__(self):
        return f"{self.user.username} - {self.date} ({self.get_status_display()})"

    @property
    def is_punched_in(self):
        return self.punch_in is not None and self.punch_out is None and self.date == timezone.localdate()

    @property
    def is_completed(self):
        return self.punch_in is not None and self.punch_out is not None

    @property
    def duration(self):
        # Shift duration is ONLY taken when both check-in and check-out happen in a day
        if self.punch_in and self.punch_out:
            return max(timedelta(0), self.punch_out - self.punch_in)
        return timedelta(0)

    @property
    def duration_minutes(self):
        return int(self.duration.total_seconds() // 60)

    @property
    def formatted_duration(self):
        if not (self.punch_in and self.punch_out):
            return "0h 0m"
        total_seconds = int(self.duration.total_seconds())
        if total_seconds <= 0:
            return "0h 0m"
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours}h {minutes}m"


class DailyTask(models.Model):
    PRIORITY_CHOICES = [
        ("low", "Low"),
        ("medium", "Medium"),
        ("high", "High"),
    ]

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("in_progress", "In Progress"),
        ("completed", "Completed"),
    ]

    TYPE_CHOICES = [
        ("dev", "Development"),
        ("support", "Support"),
        ("design", "Design"),
        ("qa", "QA / Testing"),
        ("hr", "HR & Admin"),
        ("finance", "Finance"),
    ]

    user = models.ForeignKey(
        UserModel,
        on_delete=models.CASCADE,
        related_name="daily_tasks",
    )
    title = models.CharField(max_length=255)
    description = models.TextField(blank=True, default="")
    task_type = models.CharField(
        max_length=30,
        choices=TYPE_CHOICES,
        default="dev",
        blank=True,
        null=True,
    )
    project = models.CharField(max_length=100, default="HRMS Portal", blank=True, null=True)
    assigned_by = models.CharField(max_length=100, default="Admin", blank=True, null=True)
    date = models.DateField(default=timezone.localdate)
    due_date = models.DateField(null=True, blank=True)
    priority = models.CharField(
        max_length=10,
        choices=PRIORITY_CHOICES,
        default="medium",
    )
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="pending",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-date", "-created_at"]

    def __str__(self):
        return f"{self.title} ({self.get_status_display()}) - {self.user.username}"

    @property
    def emp_id(self):
        return f"EMP-{self.user.id:04d}"

    @property
    def assignee_name(self):
        return self.user.get_full_name() or self.user.username


class LeaveRequest(models.Model):
    LEAVE_TYPES = [
        ("casual", "Casual Leave"),
        ("sick", "Sick Leave"),
        ("annual", "Annual Leave"),
        ("emergency", "Emergency Leave"),
    ]

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("approved", "Approved"),
        ("rejected", "Rejected"),
    ]

    user = models.ForeignKey(
        UserModel,
        on_delete=models.CASCADE,
        related_name="leave_requests",
    )
    leave_type = models.CharField(max_length=20, choices=LEAVE_TYPES, default="casual")
    start_date = models.DateField()
    end_date = models.DateField()
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending")
    applied_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-applied_at"]

    def __str__(self):
        return f"{self.user.username} - {self.get_leave_type_display()} ({self.get_status_display()})"

    @property
    def days_count(self):
        if self.start_date and self.end_date:
            return (self.end_date - self.start_date).days + 1
        return 1


class PermissionRequest(models.Model):
    PERMISSION_TYPES = [
        ("late_entry", "Late Entry (Morning Shift)"),
        ("early_exit", "Early Exit / Departure"),
        ("on_duty", "On Duty (OD) / Client Visit"),
        ("personal_pass", "Short Gate Pass (1-2 Hours)"),
        ("half_day_remote", "Half-Day Remote Work"),
    ]

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("approved", "Approved"),
        ("rejected", "Rejected"),
    ]

    user = models.ForeignKey(
        UserModel,
        on_delete=models.CASCADE,
        related_name="permission_requests",
    )
    permission_type = models.CharField(max_length=30, choices=PERMISSION_TYPES, default="personal_pass")
    date = models.DateField(default=timezone.localdate)
    start_time = models.TimeField(null=True, blank=True)
    end_time = models.TimeField(null=True, blank=True)
    duration_hours = models.DecimalField(max_digits=4, decimal_places=1, default=1.0)
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending")
    applied_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-applied_at"]

    def __str__(self):
        return f"{self.user.username} - {self.get_permission_type_display()} on {self.date} ({self.get_status_display()})"


class UserStatus(models.Model):
    STATUS_CHOICES = [
        ("in_office", "In Office"),
        ("remote", "Working Remote"),
        ("meeting", "In Meeting"),
        ("on_leave", "On Leave"),
        ("out_of_office", "Out of Office"),
    ]

    user = models.OneToOneField(
        UserModel,
        on_delete=models.CASCADE,
        related_name="work_status",
    )
    status = models.CharField(max_length=30, choices=STATUS_CHOICES, default="in_office")
    status_message = models.CharField(max_length=255, blank=True, default="")
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.user.username}: {self.get_status_display()}"


class EmployeeProfile(models.Model):
    user = models.OneToOneField(
        UserModel,
        on_delete=models.CASCADE,
        related_name="profile",
    )
    avatar = models.ImageField(upload_to="avatars/", null=True, blank=True)
    # Personal Info
    phone = models.CharField(max_length=30, blank=True, default="+91 98765 43210")
    personal_email = models.EmailField(blank=True, default="")
    date_of_birth = models.DateField(null=True, blank=True)
    gender = models.CharField(max_length=20, default="Not Specified")
    marital_status = models.CharField(max_length=20, default="Single")
    blood_group = models.CharField(max_length=10, default="O+")
    nationality = models.CharField(max_length=50, default="Indian")
    emergency_contact_name = models.CharField(max_length=100, blank=True, default="Emergency Contact")
    emergency_contact_phone = models.CharField(max_length=30, blank=True, default="+91 98765 00000")
    emergency_relation = models.CharField(max_length=50, blank=True, default="Family")
    current_address = models.TextField(blank=True, default="Tech Park View, Bangalore, Karnataka, India")
    permanent_address = models.TextField(blank=True, default="Tech Park View, Bangalore, Karnataka, India")

    TEAM_CHOICES = [
        ("super_admin", "Super Admin"),
        ("admin", "Admin"),
        ("hr_manager", "HR Manager"),
        ("hr_executive", "HR Executive"),
        ("manager", "Manager"),
        ("development", "Development"),
        ("finance", "Finance"),
        ("support", "Support"),
    ]

    # Employee / Organization Info
    team = models.CharField(max_length=50, choices=TEAM_CHOICES, default="development")
    department = models.CharField(max_length=100, default="Engineering")
    designation = models.CharField(max_length=100, default="Software Engineer")
    work_location = models.CharField(max_length=100, default="Headquarters (Bangalore)")
    employment_type = models.CharField(max_length=50, default="Full-Time Permanent")
    probation_status = models.CharField(max_length=50, default="Confirmed")
    reporting_manager = models.CharField(max_length=100, default="Kavitha Ramesh (CTO)")
    work_shift = models.CharField(max_length=100, default="General Shift (09:30 AM - 06:30 PM)")

    # Finance Info
    bank_name = models.CharField(max_length=100, default="HDFC Bank")
    account_holder_name = models.CharField(max_length=100, blank=True, default="")
    account_number = models.CharField(max_length=50, default="50100234567890")
    ifsc_code = models.CharField(max_length=30, default="HDFC0001234")
    branch_name = models.CharField(max_length=100, default="Electronic City Branch")
    pan_number = models.CharField(max_length=30, default="ABCDE1234F")
    uan_number = models.CharField(max_length=30, default="100987654321")
    pf_number = models.CharField(max_length=50, default="KN/BNG/12345/67890")
    payment_mode = models.CharField(max_length=50, default="Direct Bank Transfer (NEFT/RTGS)")

    # Salary Breakdown
    ctc_annual = models.DecimalField(max_digits=12, decimal_places=2, default=1200000.00)
    basic_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=50000.00)
    hra_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=25000.00)
    special_allowance_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=20000.00)
    conveyance_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=5000.00)
    pf_deduction_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=6000.00)
    tax_deduction_monthly = models.DecimalField(max_digits=10, decimal_places=2, default=4000.00)

    # Complex Histories
    increment_history = models.JSONField(default=list, blank=True)
    experience_history = models.JSONField(default=list, blank=True)
    education_history = models.JSONField(default=list, blank=True)
    documents_list = models.JSONField(default=list, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.user.username} Profile ({self.designation})"

    @property
    def gross_monthly(self):
        return self.basic_monthly + self.hra_monthly + self.special_allowance_monthly + self.conveyance_monthly

    @property
    def total_deductions_monthly(self):
        return self.pf_deduction_monthly + self.tax_deduction_monthly

    @property
    def net_monthly(self):
        return self.gross_monthly - self.total_deductions_monthly
