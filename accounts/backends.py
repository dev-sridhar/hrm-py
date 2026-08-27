from django.contrib.auth.backends import ModelBackend
from django.contrib.auth import get_user_model

UserModel = get_user_model()


class EmailAuthBackend(ModelBackend):
    """
    Custom authentication backend that allows users to log in using their email address.
    """

    def authenticate(self, request, username=None, password=None, email=None, **kwargs):
        lookup_email = email or username
        if not lookup_email or not password:
            return None

        # Look up user by email (case-insensitive)
        users = UserModel._default_manager.filter(email__iexact=lookup_email.strip())
        for user in users:
            if user.check_password(password) and self.user_can_authenticate(user):
                return user
        return None

    def get_user(self, user_id):
        try:
            user = UserModel._default_manager.get(pk=user_id)
        except UserModel.DoesNotExist:
            return None
        return user if self.user_can_authenticate(user) else None

