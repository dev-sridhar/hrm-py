from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model
from django.utils import timezone
from datetime import datetime, time, timedelta
import random

from accounts.models import Attendance, DailyTask, LeaveRequest, PermissionRequest, UserStatus, EmployeeProfile

UserModel = get_user_model()


class Command(BaseCommand):
    help = "Seeds 18+ realistic users across Super Admin, Admin, Developer, and Employee roles with rich data"

    def handle(self, *args, **options):
        self.stdout.write(self.style.NOTICE("Starting database seeding process..."))

        users_data = [
            # 1. Super Admins
            {
                "username": "superadmin",
                "email": "superadmin@company.com",
                "first_name": "Johnathan",
                "last_name": "Hayes",
                "is_superuser": True,
                "is_staff": True,
                "team": "super_admin",
                "role_desc": "Managing Director / Super Admin",
                "status": "in_office",
                "status_msg": "In Executive Office - Open door policy today",
            },
            {
                "username": "kavitha.ramesh",
                "email": "kavitha.r@company.com",
                "first_name": "Kavitha",
                "last_name": "Ramesh",
                "is_superuser": True,
                "is_staff": True,
                "team": "super_admin",
                "role_desc": "Chief Technology Officer / Super Admin",
                "status": "in_office",
                "status_msg": "Reviewing Q3 Tech Roadmap & System Architecture",
            },
            {
                "username": "Test",
                "email": "test@gmail.com",
                "first_name": "Test",
                "last_name": "Administrator",
                "is_superuser": True,
                "is_staff": True,
                "team": "super_admin",
                "role_desc": "System Administrator",
                "status": "in_office",
                "status_msg": "System operations active",
            },
            # 2. Admins
            {
                "username": "david.miller",
                "email": "david.m@company.com",
                "first_name": "David",
                "last_name": "Miller",
                "is_superuser": False,
                "is_staff": True,
                "team": "admin",
                "role_desc": "IT Infrastructure & Security Admin",
                "status": "remote",
                "status_msg": "Performing firewall and cloud access audit",
            },
            {
                "username": "arthur.morgan",
                "email": "arthur.m@company.com",
                "first_name": "Arthur",
                "last_name": "Morgan",
                "is_superuser": False,
                "is_staff": True,
                "team": "admin",
                "role_desc": "System & Network Security Admin",
                "status": "in_office",
                "status_msg": "Monitoring firewall gateways and identity endpoints",
            },
            # 3. HR Manager
            {
                "username": "marcus.vance",
                "email": "marcus.v@company.com",
                "first_name": "Marcus",
                "last_name": "Vance",
                "is_superuser": False,
                "is_staff": True,
                "team": "hr_manager",
                "role_desc": "HR Operations Director & Manager",
                "status": "in_office",
                "status_msg": "Processing monthly payroll & performance cycles",
            },
            {
                "username": "charlotte.king",
                "email": "charlotte.k@company.com",
                "first_name": "Charlotte",
                "last_name": "King",
                "is_superuser": False,
                "is_staff": True,
                "team": "hr_manager",
                "role_desc": "Senior HR Business Partner & Manager",
                "status": "in_office",
                "status_msg": "Conducting monthly leadership sync & HR strategy",
            },
            # 4. HR Executive
            {
                "username": "anita.deshmukh",
                "email": "anita.d@company.com",
                "first_name": "Anita",
                "last_name": "Deshmukh",
                "is_superuser": False,
                "is_staff": True,
                "team": "hr_executive",
                "role_desc": "Talent Acquisition & HR Executive",
                "status": "meeting",
                "status_msg": "Conducting candidate final round interviews",
            },
            {
                "username": "chloe.dupont",
                "email": "chloe.d@company.com",
                "first_name": "Chloe",
                "last_name": "Dupont",
                "is_superuser": False,
                "is_staff": True,
                "team": "hr_executive",
                "role_desc": "People Operations & Onboarding Executive",
                "status": "remote",
                "status_msg": "Coordinating orientation for new engineer batch",
            },
            # 5. Manager
            {
                "username": "sophie.laurent",
                "email": "sophie.l@company.com",
                "first_name": "Sophie",
                "last_name": "Laurent",
                "is_superuser": False,
                "is_staff": False,
                "team": "manager",
                "role_desc": "Design & UX Department Manager",
                "status": "in_office",
                "status_msg": "Creating interactive prototypes in Figma for Portal V2",
            },
            {
                "username": "vikram.singh",
                "email": "vikram.s@company.com",
                "first_name": "Vikram",
                "last_name": "Singh",
                "is_superuser": False,
                "is_staff": False,
                "team": "manager",
                "role_desc": "Operations & Growth Delivery Manager",
                "status": "on_leave",
                "status_msg": "On Annual PTO - Returning Monday",
            },
            # 6. Development
            {
                "username": "alex.chen",
                "email": "alex.chen@company.com",
                "first_name": "Alex",
                "last_name": "Chen",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Lead Full Stack Developer",
                "status": "in_office",
                "status_msg": "Refactoring Core API endpoints & GraphQL schema",
            },
            {
                "username": "priya.sharma",
                "email": "priya.sharma@company.com",
                "first_name": "Priya",
                "last_name": "Sharma",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Senior Backend Engineer (Python/Django)",
                "status": "in_office",
                "status_msg": "Optimizing database query performance & indexing",
            },
            {
                "username": "rahul.varma",
                "email": "rahul.varma@company.com",
                "first_name": "Rahul",
                "last_name": "Varma",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Senior Frontend Engineer (React/TypeScript)",
                "status": "remote",
                "status_msg": "Building Soft SaaS Design System component library",
            },
            {
                "username": "elena.rostova",
                "email": "elena.r@company.com",
                "first_name": "Elena",
                "last_name": "Rostova",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Mobile App Developer (Flutter/iOS)",
                "status": "in_office",
                "status_msg": "Testing biometric check-in module on iOS 18",
            },
            {
                "username": "siddharth.patel",
                "email": "siddharth.p@company.com",
                "first_name": "Siddharth",
                "last_name": "Patel",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "DevOps & Cloud Engineer",
                "status": "in_office",
                "status_msg": "Managing Kubernetes clusters & CI/CD deployment",
            },
            {
                "username": "clara.zhao",
                "email": "clara.z@company.com",
                "first_name": "Clara",
                "last_name": "Zhao",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Frontend Developer (UI/Web)",
                "status": "remote",
                "status_msg": "Fixing responsive mobile drawer navigation bugs",
            },
            {
                "username": "tariq.mansoor",
                "email": "tariq.m@company.com",
                "first_name": "Tariq",
                "last_name": "Mansoor",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "Database & Integration Engineer",
                "status": "meeting",
                "status_msg": "Database migration sync with backend team",
            },
            {
                "username": "arjun.nair",
                "email": "arjun.n@company.com",
                "first_name": "Arjun",
                "last_name": "Nair",
                "is_superuser": False,
                "is_staff": False,
                "team": "development",
                "role_desc": "QA & Automation Specialist",
                "status": "in_office",
                "status_msg": "Running automated Cypress cross-browser suites",
            },
            # 7. Finance
            {
                "username": "rohan.mehra",
                "email": "rohan.m@company.com",
                "first_name": "Rohan",
                "last_name": "Mehra",
                "is_superuser": False,
                "is_staff": False,
                "team": "finance",
                "role_desc": "Lead Financial Analyst & Budget Controller",
                "status": "in_office",
                "status_msg": "Analyzing quarterly financial performance & payroll forecast",
            },
            {
                "username": "rachel.green",
                "email": "rachel.g@company.com",
                "first_name": "Rachel",
                "last_name": "Green",
                "is_superuser": False,
                "is_staff": False,
                "team": "finance",
                "role_desc": "Senior Accounts & Compliance Officer",
                "status": "remote",
                "status_msg": "Reconciling GST tax filings and vendor disbursement logs",
            },
            # 8. Support
            {
                "username": "hannah.schmidt",
                "email": "hannah.s@company.com",
                "first_name": "Hannah",
                "last_name": "Schmidt",
                "is_superuser": False,
                "is_staff": False,
                "team": "support",
                "role_desc": "Customer Success & Support Lead",
                "status": "remote",
                "status_msg": "Onboarding enterprise clients & reviewing support SLA",
            },
            {
                "username": "maya.lin",
                "email": "maya.l@company.com",
                "first_name": "Maya",
                "last_name": "Lin",
                "is_superuser": False,
                "is_staff": False,
                "team": "support",
                "role_desc": "Technical Support Specialist",
                "status": "in_office",
                "status_msg": "Resolving Level 2 technical escalations & client queries",
            },
        ]

        created_users = []
        for udata in users_data:
            user, created = UserModel.objects.get_or_create(
                username=udata["username"],
                defaults={
                    "email": udata["email"],
                    "first_name": udata["first_name"],
                    "last_name": udata["last_name"],
                    "is_superuser": udata["is_superuser"],
                    "is_staff": udata["is_staff"],
                    "is_active": True,
                },
            )
            if not created:
                user.email = udata["email"]
                user.first_name = udata["first_name"]
                user.last_name = udata["last_name"]
                user.is_superuser = udata["is_superuser"]
                user.is_staff = udata["is_staff"]
                user.is_active = True
            
            user.set_password("Password123!")
            user.save()
            created_users.append(user)

            # User Status
            UserStatus.objects.update_or_create(
                user=user,
                defaults={
                    "status": udata["status"],
                    "status_message": udata["status_msg"],
                },
            )

            # Employee Profile Details
            if udata["is_superuser"]:
                dept = "Executive Management"
                desig = udata.get("role_desc", "Managing Director")
                ctc = 2800000.00
                basic = 110000.00
                hra = 55000.00
                special = 45000.00
                conveyance = 10000.00
                pf = 12000.00
                tax = 15000.00
                manager = "Board of Directors"
                loc = "Executive Headquarters (Bangalore)"
            elif udata["is_staff"]:
                dept = "Human Resources & Operations"
                desig = udata.get("role_desc", "HR Operations Lead")
                ctc = 1650000.00
                basic = 65000.00
                hra = 32500.00
                special = 25000.00
                conveyance = 8000.00
                pf = 7800.00
                tax = 6500.00
                manager = "Johnathan Hayes (Managing Director)"
                loc = "Tech Park HQ (Bangalore)"
            elif "dev" in udata["username"] or udata["username"] in {"alex.chen", "priya.sharma", "rahul.varma", "elena.rostova", "siddharth.patel", "clara.zhao", "tariq.mansoor"}:
                dept = "Engineering & Product"
                desig = udata.get("role_desc", "Senior Full-Stack Engineer")
                ctc = 1450000.00
                basic = 60000.00
                hra = 30000.00
                special = 20000.00
                conveyance = 6000.00
                pf = 7200.00
                tax = 5500.00
                manager = "Kavitha Ramesh (CTO)"
                loc = "Innovation Hub (Bangalore)"
            else:
                dept = "Design & Quality Assurance"
                desig = udata.get("role_desc", "UI/UX Specialist")
                ctc = 1050000.00
                basic = 45000.00
                hra = 22500.00
                special = 14000.00
                conveyance = 5000.00
                pf = 5400.00
                tax = 3500.00
                manager = "Marcus Vance (HR Operations Director)"
                loc = "Tech Park HQ (Bangalore)"

            inc_logs = [
                {
                    "effective_date": "Apr 01, 2026",
                    "previous_ctc": f"₹{(ctc * 0.85):,.0f}",
                    "new_ctc": f"₹{ctc:,.0f}",
                    "increment_pct": "+17.6%",
                    "title_change": desig,
                    "reviewer": "Kavitha Ramesh (CTO)",
                    "rating": "Outstanding (4.9 / 5.0)",
                    "notes": "Exceptional leadership in modular SaaS portal architecture and cross-team delivery.",
                },
                {
                    "effective_date": "Apr 01, 2025",
                    "previous_ctc": f"₹{(ctc * 0.70):,.0f}",
                    "new_ctc": f"₹{(ctc * 0.85):,.0f}",
                    "increment_pct": "+21.4%",
                    "title_change": "Senior Associate",
                    "reviewer": "Johnathan Hayes (MD)",
                    "rating": "Exceeds Expectations (4.7 / 5.0)",
                    "notes": "Strong performance during annual roadmap rollout and zero downtime deployment.",
                },
                {
                    "effective_date": "Oct 01, 2024",
                    "previous_ctc": f"₹{(ctc * 0.60):,.0f}",
                    "new_ctc": f"₹{(ctc * 0.70):,.0f}",
                    "increment_pct": "+16.6%",
                    "title_change": "Associate Level II",
                    "reviewer": "Marcus Vance",
                    "rating": "Meets & Exceeds Expectations (4.5 / 5.0)",
                    "notes": "Consistent delivery and proactive ownership of platform reliability.",
                },
            ]

            exp_logs = [
                {
                    "company": "Infosys Technologies Ltd",
                    "role": "Software Engineer II",
                    "duration": "Jul 2021 - Aug 2024 (3 Years 2 Mos)",
                    "description": "Architected high-throughput microservices handling 2M+ requests daily with 99.99% uptime.",
                },
                {
                    "company": "Cognizant Solutions",
                    "role": "Junior Systems Developer",
                    "duration": "Jun 2019 - Jun 2021 (2 Years)",
                    "description": "Implemented full-stack client dashboards, REST APIs, and automated CI/CD deployment pipelines.",
                },
            ]

            edu_logs = [
                {
                    "degree": "Bachelor of Technology (B.Tech) in Computer Science",
                    "institution": "National Institute of Technology (NIT)",
                    "year": "2015 - 2019",
                    "grade": "First Class with Distinction (8.8 CGPA)",
                },
                {
                    "degree": "AWS Certified Solutions Architect - Associate",
                    "institution": "Amazon Web Services (AWS)",
                    "year": "Certified 2023 - 2026",
                    "grade": "Credential ID: AWS-ARC-987410",
                },
            ]

            docs_list = [
                {
                    "title": "Government Identity Proof (Passport / Aadhar)",
                    "doc_type": "Identity Verification",
                    "uploaded_date": "Aug 20, 2024",
                    "status": "Verified",
                    "file_size": "2.4 MB",
                },
                {
                    "title": "Signed Employment Offer Letter & Agreement",
                    "doc_type": "HR Contract",
                    "uploaded_date": "Aug 22, 2024",
                    "status": "Verified",
                    "file_size": "1.8 MB",
                },
                {
                    "title": "Degree & Academic Graduation Certificate",
                    "doc_type": "Education",
                    "uploaded_date": "Aug 23, 2024",
                    "status": "Verified",
                    "file_size": "3.1 MB",
                },
                {
                    "title": "Previous Organization Relieving & Service Letter",
                    "doc_type": "Experience",
                    "uploaded_date": "Aug 24, 2024",
                    "status": "Verified",
                    "file_size": "1.5 MB",
                },
                {
                    "title": "Non-Disclosure & Intellectual Property Agreement",
                    "doc_type": "Legal / NDA",
                    "uploaded_date": "Aug 24, 2024",
                    "status": "Active",
                    "file_size": "1.2 MB",
                },
            ]

            EmployeeProfile.objects.update_or_create(
                user=user,
                defaults={
                    "phone": f"+91 98450 {user.id:05d}",
                    "personal_email": f"{user.username}.personal@gmail.com",
                    "date_of_birth": datetime(1992, (user.id % 12) + 1, (user.id % 25) + 1).date(),
                    "gender": "Male" if user.id % 2 == 1 else "Female",
                    "marital_status": "Married" if user.id % 3 == 0 else "Single",
                    "blood_group": ["O+", "A+", "B+", "AB+"][user.id % 4],
                    "nationality": "Indian",
                    "emergency_contact_name": f"{user.first_name}'s Family Contact",
                    "emergency_contact_phone": f"+91 99000 {user.id:05d}",
                    "emergency_relation": "Spouse" if user.id % 3 == 0 else "Parent",
                    "current_address": f"Flat {100 + user.id}, Tech Park Enclave, Outer Ring Road, Bangalore - 560103",
                    "permanent_address": f"House #{40 + user.id}, Palm Meadows, Green Avenue, Bangalore - 560066",
                    "team": udata.get("team", "development"),
                    "department": dept,
                    "designation": desig,
                    "work_location": loc,
                    "employment_type": "Full-Time Permanent",
                    "probation_status": "Confirmed",
                    "reporting_manager": manager,
                    "work_shift": "General Shift (09:30 AM - 06:30 PM)",
                    "bank_name": ["HDFC Bank", "ICICI Bank", "State Bank of India", "Axis Bank"][user.id % 4],
                    "account_holder_name": user.get_full_name() or user.username,
                    "account_number": f"50100{user.id:04d}890123",
                    "ifsc_code": "HDFC0001234",
                    "branch_name": "Electronic City Tech Zone Branch",
                    "pan_number": f"ABCDE{user.id:04d}F",
                    "uan_number": f"10098{user.id:05d}321",
                    "pf_number": f"KN/BNG/{user.id:05d}/67890",
                    "payment_mode": "Direct Bank Transfer (NEFT/RTGS)",
                    "ctc_annual": ctc,
                    "basic_monthly": basic,
                    "hra_monthly": hra,
                    "special_allowance_monthly": special,
                    "conveyance_monthly": conveyance,
                    "pf_deduction_monthly": pf,
                    "tax_deduction_monthly": tax,
                    "increment_history": inc_logs,
                    "experience_history": exp_logs,
                    "education_history": edu_logs,
                    "documents_list": docs_list,
                    "avatar": f"avatars/{user.username}.jpg",
                },
            )

        self.stdout.write(self.style.SUCCESS(f"Successfully seeded {len(created_users)} users, profiles & presence statuses!"))

        # Seed Attendances
        today = timezone.localdate()
        days_to_seed = 10

        self.stdout.write(self.style.NOTICE("Seeding historical attendance records..."))
        for user in created_users:
            for d in range(days_to_seed, -1, -1):
                att_date = today - timedelta(days=d)
                
                # Skip Sundays
                if att_date.weekday() == 6:
                    continue

                existing = Attendance.objects.filter(user=user, date=att_date).first()
                if not existing:
                    # Determine status
                    if d == 0:
                        # Today's status based on user's current status
                        u_stat = user.work_status.status if hasattr(user, 'work_status') else "in_office"
                        if u_stat == "on_leave":
                            status = "leave"
                            p_in = None
                            p_out = None
                        elif u_stat in ["in_office", "remote", "meeting"]:
                            status = "present"
                            # Checked in around 9:00 - 9:30 AM today
                            p_in = timezone.make_aware(datetime.combine(att_date, time(random.randint(8, 9), random.randint(15, 55))))
                            p_out = None  # currently punched in
                        else:
                            status = "present"
                            p_in = timezone.make_aware(datetime.combine(att_date, time(9, 15)))
                            p_out = None
                    else:
                        # Past days: mostly present, occasionally leave or absent
                        rand_val = random.random()
                        if rand_val < 0.85:
                            status = "present"
                            start_hour = random.randint(8, 9)
                            start_min = random.randint(15, 55)
                            p_in = timezone.make_aware(datetime.combine(att_date, time(start_hour, start_min)))
                            
                            end_hour = random.randint(17, 19)
                            end_min = random.randint(10, 50)
                            p_out = timezone.make_aware(datetime.combine(att_date, time(end_hour, end_min)))
                        elif rand_val < 0.95:
                            status = "leave"
                            p_in = None
                            p_out = None
                        else:
                            status = "lop"
                            p_in = None
                            p_out = None

                    Attendance.objects.create(
                        user=user,
                        date=att_date,
                        punch_in=p_in,
                        punch_out=p_out,
                        status=status,
                        notes="Regular biometric log" if status == "present" else "Approved absence",
                    )

        self.stdout.write(self.style.SUCCESS("Attendance records seeded successfully!"))

        # Seed Daily Tasks
        self.stdout.write(self.style.NOTICE("Seeding rich daily tasks..."))
        tasks_pool = [
            ("Optimize database indexing on Attendance and Leave queries", "Run EXPLAIN ANALYZE on complex filter queries to reduce latency under 20ms.", "high", "in_progress"),
            ("Implement OAuth2 Single Sign-On integration", "Allow enterprise employees to log in using corporate Google/Okta credentials.", "high", "pending"),
            ("Design Soft SaaS Mobile App UI specifications in Figma", "Prepare component tokens, color variables, and interactive mobile prototype.", "medium", "completed"),
            ("Build automated end-to-end Cypress tests", "Cover login, punch in/out, task creation, and leave approval workflows.", "medium", "in_progress"),
            ("Conduct weekly sprint review and backlog refinement", "Align product roadmap priorities with engineering leads.", "medium", "completed"),
            ("Configure CloudWatch logging and Sentry error monitoring", "Ensure zero untracked exceptions across production endpoints.", "high", "completed"),
            ("Review and approve pending employee reimbursement receipts", "Verify travel expense documentation for quarterly client meetings.", "low", "pending"),
            ("Refactor dashboard CSS into modular design tokens", "Replace legacy styles with clean CSS variables and soft card shadows.", "medium", "completed"),
            ("Set up Redis caching layer for active presence board", "Cache live presence broadcast pings to reduce database read load.", "high", "pending"),
            ("Conduct team 1-on-1 performance and feedback check-ins", "Review individual milestones and professional development goals.", "low", "in_progress"),
            ("Audit employee permission roles and access levels", "Ensure least-privilege access across production database resources.", "high", "pending"),
            ("Draft release notes for Employee Portal Version 2.4", "Summarize newly released attendance timesheet and task management features.", "low", "completed"),
        ]

        for i, user in enumerate(created_users):
            # Give 2-3 tasks to each user
            user_tasks = tasks_pool[i % len(tasks_pool): (i % len(tasks_pool)) + 2]
            for title, desc, prio, stat in user_tasks:
                DailyTask.objects.get_or_create(
                    user=user,
                    title=f"{title} - {user.first_name}",
                    defaults={
                        "description": desc,
                        "priority": prio,
                        "status": stat,
                        "due_date": today + timedelta(days=random.randint(1, 14)),
                    },
                )

        self.stdout.write(self.style.SUCCESS("Daily tasks seeded successfully!"))

        # Seed Leave Requests
        self.stdout.write(self.style.NOTICE("Seeding leave requests..."))
        leave_samples = [
            ("casual", 1, "Personal family commitment and home errands", "approved"),
            ("sick", 2, "Viral flu and doctor recommended bed rest", "approved"),
            ("annual", 4, "Annual family vacation and travel", "approved"),
            ("emergency", 1, "Urgent home plumbing repair and maintenance", "pending"),
            ("casual", 2, "Attending cousin wedding ceremony in native town", "pending"),
            ("sick", 1, "Dental checkup and routine appointment", "rejected"),
            ("annual", 5, "Year-end holiday travel with family", "pending"),
            ("casual", 1, "Personal vehicle registration renewal appointment", "approved"),
        ]

        for i, user in enumerate(created_users[:12]):
            l_type, days_len, reason, l_stat = leave_samples[i % len(leave_samples)]
            start_d = today + timedelta(days=random.randint(-10, 20))
            end_d = start_d + timedelta(days=days_len - 1)
            
            LeaveRequest.objects.get_or_create(
                user=user,
                start_date=start_d,
                end_date=end_d,
                defaults={
                    "leave_type": l_type,
                    "reason": reason,
                    "status": l_stat,
                },
            )

        self.stdout.write(self.style.SUCCESS("Leave requests seeded successfully!"))

        # Seed Permission Requests & Gate Passes
        self.stdout.write(self.style.NOTICE("Seeding permissions and gate passes..."))
        permission_samples = [
            ("late_entry", time(9, 30), time(11, 0), 1.5, "Traffic delay on outer ring road due to road construction", "approved"),
            ("early_exit", time(16, 30), time(18, 30), 2.0, "Family medical appointment and hospital consultation", "pending"),
            ("on_duty", time(14, 0), time(17, 0), 3.0, "Client on-site deployment review at Infosys campus", "approved"),
            ("personal_pass", time(12, 30), time(14, 0), 1.5, "Bank document verification for loan paperwork", "approved"),
            ("half_day_remote", time(14, 0), time(18, 30), 4.0, "Residential electrical maintenance outage", "pending"),
            ("late_entry", time(9, 30), time(10, 30), 1.0, "Metro rail signal interruption delay", "approved"),
            ("on_duty", time(10, 0), time(13, 0), 3.0, "Vendor contract finalization and security inspection", "approved"),
            ("early_exit", time(17, 0), time(18, 30), 1.5, "Attending continuing education certification exam", "rejected"),
        ]

        for i, user in enumerate(created_users):
            p_type, s_time, e_time, p_hours, p_reason, p_stat = permission_samples[i % len(permission_samples)]
            p_date = today + timedelta(days=random.randint(-7, 7))

            PermissionRequest.objects.get_or_create(
                user=user,
                date=p_date,
                permission_type=p_type,
                defaults={
                    "start_time": s_time,
                    "end_time": e_time,
                    "duration_hours": p_hours,
                    "reason": p_reason,
                    "status": p_stat,
                },
            )

        self.stdout.write(self.style.SUCCESS("Permissions and gate passes seeded successfully!"))
        self.stdout.write(self.style.SUCCESS("=== SEEDING COMPLETED SUCCESSFULLY ==="))
        self.stdout.write(self.style.SUCCESS(f"Total Users in DB: {UserModel.objects.count()}"))
        self.stdout.write(self.style.SUCCESS(f"Total Attendance Records: {Attendance.objects.count()}"))
        self.stdout.write(self.style.SUCCESS(f"Total Tasks: {DailyTask.objects.count()}"))
        self.stdout.write(self.style.SUCCESS(f"Total Leave Requests: {LeaveRequest.objects.count()}"))
        self.stdout.write(self.style.SUCCESS(f"Total Permission Requests: {PermissionRequest.objects.count()}"))


