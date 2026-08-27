from django import forms
from django.contrib.auth import authenticate, get_user_model
from django.core.exceptions import ValidationError
from .models import DailyTask, Attendance, LeaveRequest, PermissionRequest

UserModel = get_user_model()


class EmailLoginForm(forms.Form):
    email = forms.EmailField(
        label="Email Address",
        widget=forms.EmailInput(
            attrs={
                "class": "form-input",
                "placeholder": "name@example.com",
                "autocomplete": "email",
                "autofocus": True,
            }
        ),
    )
    password = forms.CharField(
        label="Password",
        widget=forms.PasswordInput(
            attrs={
                "class": "form-input",
                "placeholder": "Enter your password",
                "autocomplete": "current-password",
            }
        ),
    )
    remember_me = forms.BooleanField(
        required=False,
        initial=False,
        label="Remember me",
        widget=forms.CheckboxInput(
            attrs={
                "class": "form-checkbox",
            }
        ),
    )

    def __init__(self, request=None, *args, **kwargs):
        self.request = request
        self.user_cache = None
        super().__init__(*args, **kwargs)

    def clean(self):
        cleaned_data = super().clean()
        email = cleaned_data.get("email")
        password = cleaned_data.get("password")

        if email and password:
            self.user_cache = authenticate(
                self.request,
                username=email,
                email=email,
                password=password,
            )
            if self.user_cache is None:
                raise ValidationError(
                    "Invalid email address or password. Please check your credentials and try again.",
                    code="invalid_login",
                )
            elif not self.user_cache.is_active:
                raise ValidationError(
                    "This account is currently inactive. Please contact support.",
                    code="inactive",
                )
        return cleaned_data

    def get_user(self):
        return self.user_cache


class RegisterForm(forms.Form):
    username = forms.CharField(
        max_length=150,
        label="Username",
        widget=forms.TextInput(
            attrs={
                "class": "form-input",
                "placeholder": "Choose a username",
                "autocomplete": "username",
            }
        ),
    )
    email = forms.EmailField(
        label="Email Address",
        widget=forms.EmailInput(
            attrs={
                "class": "form-input",
                "placeholder": "name@example.com",
                "autocomplete": "email",
            }
        ),
    )
    password = forms.CharField(
        label="Password",
        widget=forms.PasswordInput(
            attrs={
                "class": "form-input",
                "placeholder": "Create a strong password",
                "autocomplete": "new-password",
            }
        ),
    )
    confirm_password = forms.CharField(
        label="Confirm Password",
        widget=forms.PasswordInput(
            attrs={
                "class": "form-input",
                "placeholder": "Repeat your password",
                "autocomplete": "new-password",
            }
        ),
    )

    def clean_username(self):
        username = self.cleaned_data.get("username", "").strip()
        if UserModel.objects.filter(username__iexact=username).exists():
            raise ValidationError("A user with this username already exists.")
        return username

    def clean_email(self):
        email = self.cleaned_data.get("email", "").strip().lower()
        if UserModel.objects.filter(email__iexact=email).exists():
            raise ValidationError("A user with this email address already exists.")
        return email

    def clean(self):
        cleaned_data = super().clean()
        password = cleaned_data.get("password")
        confirm_password = cleaned_data.get("confirm_password")

        if password and confirm_password and password != confirm_password:
            self.add_error("confirm_password", "Passwords do not match.")
        return cleaned_data

    def save(self):
        username = self.cleaned_data["username"]
        email = self.cleaned_data["email"]
        password = self.cleaned_data["password"]
        user = UserModel.objects.create_user(
            username=username,
            email=email,
            password=password,
        )
        return user


class TaskForm(forms.ModelForm):
    class Meta:
        model = DailyTask
        fields = ["title", "description", "priority", "due_date", "status"]
        widgets = {
            "title": forms.TextInput(
                attrs={
                    "class": "dash-form-input",
                    "placeholder": "e.g., Complete UI integration for auth module",
                    "required": True,
                }
            ),
            "description": forms.Textarea(
                attrs={
                    "class": "dash-form-textarea",
                    "placeholder": "Add any notes or specifications...",
                    "rows": 3,
                }
            ),
            "priority": forms.Select(
                attrs={"class": "dash-form-select"}
            ),
            "status": forms.Select(
                attrs={"class": "dash-form-select"}
            ),
            "due_date": forms.DateInput(
                attrs={"class": "dash-form-input", "type": "date"}
            ),
        }


class LeaveRequestForm(forms.ModelForm):
    class Meta:
        model = LeaveRequest
        fields = ["leave_type", "start_date", "end_date", "reason"]
        widgets = {
            "leave_type": forms.Select(attrs={"class": "dash-form-select"}),
            "start_date": forms.DateInput(attrs={"class": "dash-form-input", "type": "date", "required": True}),
            "end_date": forms.DateInput(attrs={"class": "dash-form-input", "type": "date", "required": True}),
            "reason": forms.Textarea(attrs={"class": "dash-form-textarea", "placeholder": "Reason for leave...", "rows": 3, "required": True}),
        }


class PermissionRequestForm(forms.ModelForm):
    class Meta:
        model = PermissionRequest
        fields = ["permission_type", "date", "start_time", "end_time", "duration_hours", "reason"]
        widgets = {
            "permission_type": forms.Select(attrs={"class": "dash-form-select"}),
            "date": forms.DateInput(attrs={"class": "dash-form-input", "type": "date", "required": True}),
            "start_time": forms.TimeInput(attrs={"class": "dash-form-input", "type": "time"}),
            "end_time": forms.TimeInput(attrs={"class": "dash-form-input", "type": "time"}),
            "duration_hours": forms.NumberInput(attrs={"class": "dash-form-input", "step": "0.5", "min": "0.5", "max": "8.0", "required": True}),
            "reason": forms.Textarea(attrs={"class": "dash-form-textarea", "placeholder": "Specify exact reason, client meeting details, or emergency notes...", "rows": 3, "required": True}),
        }



class ProfileUpdateForm(forms.Form):
    first_name = forms.CharField(
        max_length=150,
        required=False,
        label="First Name",
        widget=forms.TextInput(attrs={"class": "dash-form-input", "placeholder": "First Name"}),
    )
    last_name = forms.CharField(
        max_length=150,
        required=False,
        label="Last Name",
        widget=forms.TextInput(attrs={"class": "dash-form-input", "placeholder": "Last Name"}),
    )
    email = forms.EmailField(
        label="Email Address",
        widget=forms.EmailInput(attrs={"class": "dash-form-input", "placeholder": "name@example.com"}),
    )

    def __init__(self, user=None, *args, **kwargs):
        self.user = user
        super().__init__(*args, **kwargs)
        if self.user:
            self.fields["first_name"].initial = self.user.first_name
            self.fields["last_name"].initial = self.user.last_name
            self.fields["email"].initial = self.user.email

    def clean_email(self):
        email = self.cleaned_data.get("email", "").strip().lower()
        if UserModel.objects.filter(email__iexact=email).exclude(pk=self.user.pk).exists():
            raise ValidationError("A user with this email address already exists.")
        return email

    def save(self):
        self.user.first_name = self.cleaned_data["first_name"]
        self.user.last_name = self.cleaned_data["last_name"]
        self.user.email = self.cleaned_data["email"]
        self.user.save()
        return self.user


class CustomPasswordChangeForm(forms.Form):
    old_password = forms.CharField(
        label="Current Password",
        widget=forms.PasswordInput(attrs={"class": "dash-form-input", "placeholder": "Enter current password"}),
    )
    new_password1 = forms.CharField(
        label="New Password",
        widget=forms.PasswordInput(attrs={"class": "dash-form-input", "placeholder": "Enter new password"}),
    )
    new_password2 = forms.CharField(
        label="Confirm New Password",
        widget=forms.PasswordInput(attrs={"class": "dash-form-input", "placeholder": "Confirm new password"}),
    )

    def __init__(self, user=None, *args, **kwargs):
        self.user = user
        super().__init__(*args, **kwargs)

    def clean_old_password(self):
        old_password = self.cleaned_data.get("old_password")
        if not self.user.check_password(old_password):
            raise ValidationError("Current password is not correct.")
        return old_password

    def clean(self):
        cleaned_data = super().clean()
        new_password1 = cleaned_data.get("new_password1")
        new_password2 = cleaned_data.get("new_password2")

        if new_password1 and new_password2 and new_password1 != new_password2:
            self.add_error("new_password2", "New passwords do not match.")
        return cleaned_data

    def save(self):
        self.user.set_password(self.cleaned_data["new_password1"])
        self.user.save()
        return self.user

