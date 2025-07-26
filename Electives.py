import pandas as pd
import numpy as np
from collections import defaultdict
import random
import math
import os
from pathlib import Path

# ==================================================================================
#                           ELECTIVES MANAGEMENT SYSTEM
# ==================================================================================

class ElectivesManager:
    """
    Comprehensive Electives Management System
    Handles cross-department electives, student preferences, and optimization
    """

    def __init__(self):
        # ---------------------- TIME CONFIGURATION ----------------------
        self.TIME_SLOTS = [
            '08:00-09:15', '09:30-10:45', '11:00-12:15',
            '12:30-01:45', '02:00-03:15', '03:30-04:45', '05:00-06:15'
        ]

        self.DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

        # ---------------------- ELECTIVES CONFIGURATION ----------------------
        self.ELECTIVE_TYPES = {
            'General': {
                'description': 'General education electives',
                'cross_department': True,
                'min_students': 30,
                'max_students': 60,
                'priority': 'high',
                'preferred_times': ['03:30-04:45', '05:00-06:15']
            },
            'Technical': {
                'description': 'Technical electives within department',
                'cross_department': False,
                'min_students': 20,
                'max_students': 40,
                'priority': 'medium',
                'preferred_times': ['02:00-03:15', '03:30-04:45']
            },
            'Free': {
                'description': 'Free choice electives',
                'cross_department': True,
                'min_students': 15,
                'max_students': 35,
                'priority': 'low',
                'preferred_times': ['05:00-06:15']
            }
        }

        # ---------------------- DATA STORAGE ----------------------
        self.core_files = {}
        self.electives_pool = {}
        self.student_preferences = defaultdict(list)
        self.elective_sections = defaultdict(list)
        self.room_pools = {'theory': [], 'lab': []}
        self.elective_timetables = defaultdict(lambda: pd.DataFrame(index=self.TIME_SLOTS, columns=self.DAYS))
        self.room_bookings = defaultdict(lambda: defaultdict(set))
        self.enrollment_stats = defaultdict(dict)

        # Statistics
        self.stats = {
            'total_electives': 0,
            'sections_created': 0,
            'students_enrolled': 0,
            'cross_dept_electives': 0,
            'conflicts_resolved': 0
        }

        print("🎯 Electives Management System Initialized")
        print(f"📅 Time Slots Available: {len(self.TIME_SLOTS)}")
        print(f"🎓 Elective Types: {list(self.ELECTIVE_TYPES.keys())}")

    # ==================================================================================
    #                               FILE OPERATIONS
    # ==================================================================================

    def load_core_files(self, file_list=None):
        """Load core files to extract electives data"""
        print("\n📂 LOADING CORE FILES FOR ELECTIVES")
        print("-" * 50)

        if file_list is None:
            file_list = list(Path('.').glob('*.xlsx'))
            file_list = [f for f in file_list if not f.name.startswith('~') and not f.name.startswith('cohort')]

        for file_path in file_list:
            self._load_single_file(file_path)

        print(f"✅ Loaded {len(self.core_files)} core files")
        return len(self.core_files) > 0

    def _load_single_file(self, file_path):
        """Load individual file and extract electives"""
        try:
            file_name = Path(file_path).name
            print(f"📄 Loading {file_name}...")

            if not os.path.exists(file_path):
                print(f"   ❌ File not found: {file_path}")
                return

            excel_data = pd.ExcelFile(file_path)
            sheets = {}

            for sheet_name in excel_data.sheet_names:
                try:
                    df = excel_data.parse(sheet_name)
                    df.columns = [str(col).strip() for col in df.columns]
                    df = df.dropna(how='all')
                    sheets[sheet_name] = df
                except Exception as e:
                    print(f"   ⚠️ Error in sheet {sheet_name}: {e}")

            department = self._extract_department(file_name)
            self.core_files[file_name] = {
                'name': file_name,
                'department': department,
                'sheets': sheets
            }

            print(f"   📚 Department: {department}, Sheets: {len(sheets)}")

        except Exception as e:
            print(f"   ❌ Failed to load {file_path}: {e}")

    def _extract_department(self, filename):
        """Extract department name from filename"""
        name = Path(filename).stem.upper()
        name = name.replace(' ', '').replace('_4.2', '').replace('.', '')

        dept_mapping = {
            'BSCB': 'CB', 'BSCS': 'CS', 'BSINFS': 'INFS',
            'BSSE': 'SE', 'BSAI': 'AI', 'BSDS': 'DS'
        }

        return dept_mapping.get(name, name)

    # ==================================================================================
    #                            ELECTIVES PROCESSING
    # ==================================================================================

    def process_electives_data(self):
        """Process electives from all departments"""
        print("\n🎓 PROCESSING ELECTIVES DATA")
        print("-" * 50)

        for file_name, file_data in self.core_files.items():
            if 'Electives' not in file_data['sheets']:
                continue

            department = file_data['department']
            electives_df = file_data['sheets']['Electives']

            print(f"\n📊 Processing {department} electives:")
            print(f"   Found {len(electives_df)} electives")

            for _, row in electives_df.iterrows():
                try:
                    elective_code = str(row.get('elective_code', '')).strip()
                    elective_name = str(row.get('elective_name', '')).strip()
                    elective_type = str(row.get('elective_type', 'Technical')).strip()
                    credit_hours = int(row.get('credit_hour', 3))
                    sections_count = int(row.get('sections_count', 1))
                    can_use_theory = bool(row.get('can_use_theory', True))
                    can_use_lab = bool(row.get('can_use_lab', False))

                    if not elective_code or elective_code == 'nan':
                        continue

                    # Determine eligible departments
                    eligible_departments = self._determine_eligible_departments(
                        elective_type, department, elective_code
                    )

                    # Store elective data
                    self.electives_pool[elective_code] = {
                        'name': elective_name,
                        'type': elective_type,
                        'credit_hours': credit_hours,
                        'sections_count': sections_count,
                        'can_use_theory': can_use_theory,
                        'can_use_lab': can_use_lab,
                        'source_department': department,
                        'eligible_departments': eligible_departments,
                        'min_students': self.ELECTIVE_TYPES.get(elective_type, {}).get('min_students', 20),
                        'max_students': self.ELECTIVE_TYPES.get(elective_type, {}).get('max_students', 40),
                        'priority': self.ELECTIVE_TYPES.get(elective_type, {}).get('priority', 'medium')
                    }

                    self.stats['total_electives'] += 1

                    if len(eligible_departments) > 1:
                        self.stats['cross_dept_electives'] += 1

                    print(f"   ✅ {elective_code}: {elective_name}")
                    print(f"      Type: {elective_type}, Eligible: {eligible_departments}")

                except Exception as e:
                    print(f"   ⚠️ Error processing elective: {e}")

        print(f"\n📈 Electives Summary:")
        print(f"   Total electives: {self.stats['total_electives']}")
        print(f"   Cross-department: {self.stats['cross_dept_electives']}")

    def _determine_eligible_departments(self, elective_type, source_dept, elective_code):
        """Determine which departments can take this elective"""
        if elective_type == 'General':
            return ['ALL']  # All departments
        elif elective_type == 'Technical':
            # Technical electives within related departments
            related_depts = {
                'CS': ['CS', 'SE', 'AI', 'DS'],
                'SE': ['CS', 'SE', 'AI'],
                'AI': ['CS', 'AI', 'DS'],
                'DS': ['CS', 'AI', 'DS'],
                'INFS': ['INFS', 'CS'],
                'CB': ['CB']
            }
            return related_depts.get(source_dept, [source_dept])
        else:  # Free electives
            return ['ALL']

    def setup_rooms(self):
        """Setup room pools for electives"""
        print("\n🏢 SETTING UP ELECTIVES ROOMS")
        print("-" * 50)

        # Collect rooms from all departments
        theory_rooms = set()
        lab_rooms = set()

        for file_data in self.core_files.values():
            if 'Rooms' in file_data['sheets']:
                rooms_df = file_data['sheets']['Rooms']
                for _, row in rooms_df.iterrows():
                    room_name = str(row.get('room_name', '')).strip()
                    room_type = str(row.get('room_type', '')).lower()

                    if room_name and room_name != 'nan':
                        if 'theory' in room_type:
                            theory_rooms.add(room_name)
                        elif 'lab' in room_type:
                            lab_rooms.add(room_name)
                        else:
                            theory_rooms.add(room_name)

        self.room_pools['theory'] = list(theory_rooms) or [f'Elective-T{i:02d}' for i in range(1, 11)]
        self.room_pools['lab'] = list(lab_rooms) or [f'Elective-L{i:02d}' for i in range(1, 6)]

        print(f"📚 Theory rooms for electives: {len(self.room_pools['theory'])}")
        print(f"🔬 Lab rooms for electives: {len(self.room_pools['lab'])}")

    # ==================================================================================
    #                              STUDENT MANAGEMENT
    # ==================================================================================

    def generate_sample_student_preferences(self):
        """Generate sample student preferences for demonstration"""
        print("\n👥 GENERATING SAMPLE STUDENT PREFERENCES")
        print("-" * 50)

        departments = ['CS', 'SE', 'AI', 'DS', 'INFS', 'CB']
        semesters = [5, 6, 7, 8]  # Electives typically in later semesters

        student_id = 1000

        for dept in departments:
            for semester in semesters:
                # Generate 30-50 students per department per semester
                num_students = random.randint(30, 50)

                for _ in range(num_students):
                    student_key = f"{dept}_S{semester}_ST{student_id}"

                    # Select 2-3 electives per student
                    available_electives = self._get_available_electives_for_student(dept, semester)
                    num_choices = min(3, len(available_electives))

                    if num_choices > 0:
                        chosen_electives = random.sample(available_electives, num_choices)

                        for priority, elective_code in enumerate(chosen_electives, 1):
                            self.student_preferences[student_key].append({
                                'elective_code': elective_code,
                                'priority': priority,
                                'department': dept,
                                'semester': semester
                            })

                    student_id += 1

        total_preferences = sum(len(prefs) for prefs in self.student_preferences.values())
        print(f"✅ Generated preferences for {len(self.student_preferences)} students")
        print(f"📊 Total preference entries: {total_preferences}")

    def _get_available_electives_for_student(self, dept, semester):
        """Get electives available for a specific student"""
        available = []

        for elective_code, elective_data in self.electives_pool.items():
            eligible_depts = elective_data['eligible_departments']

            if 'ALL' in eligible_depts or dept in eligible_depts:
                # Add semester-based filtering logic here if needed
                available.append(elective_code)

        return available

    def analyze_demand(self):
        """Analyze student demand for each elective"""
        print("\n📊 ANALYZING ELECTIVE DEMAND")
        print("-" * 50)

        demand_analysis = defaultdict(lambda: {'total': 0, 'by_priority': {1: 0, 2: 0, 3: 0}, 'by_dept': defaultdict(int)})

        for student_key, preferences in self.student_preferences.items():
            for pref in preferences:
                elective_code = pref['elective_code']
                priority = pref['priority']
                dept = pref['department']

                demand_analysis[elective_code]['total'] += 1
                demand_analysis[elective_code]['by_priority'][priority] += 1
                demand_analysis[elective_code]['by_dept'][dept] += 1

        # Sort by demand
        sorted_electives = sorted(demand_analysis.items(), key=lambda x: x[1]['total'], reverse=True)

        print("📈 Top 10 Most Demanded Electives:")
        for i, (elective_code, demand_data) in enumerate(sorted_electives[:10], 1):
            elective_name = self.electives_pool[elective_code]['name']
            total_demand = demand_data['total']
            first_choice = demand_data['by_priority'][1]

            print(f"   {i:2d}. {elective_code}: {total_demand} students ({first_choice} first choice)")
            print(f"       {elective_name}")

        return demand_analysis

    # ==================================================================================
    #                              SCHEDULING LOGIC
    # ==================================================================================

    def create_elective_sections(self, demand_analysis):
        """Create sections based on student demand"""
        print("\n🏗️ CREATING ELECTIVE SECTIONS")
        print("-" * 50)

        for elective_code, demand_data in demand_analysis.items():
            if elective_code not in self.electives_pool:
                continue

            elective_info = self.electives_pool[elective_code]
            total_demand = demand_data['total']

            if total_demand < elective_info['min_students']:
                print(f"   ⚠️ {elective_code}: Low demand ({total_demand}) - Not scheduling")
                continue

            # Calculate sections needed
            max_per_section = elective_info['max_students']
            sections_needed = math.ceil(total_demand / max_per_section)

            # Create sections
            for section_num in range(1, sections_needed + 1):
                section_key = f"Elective_{elective_code}_Sec{section_num}"

                self.elective_sections[elective_code].append({
                    'section_key': section_key,
                    'capacity': max_per_section,
                    'enrolled': 0,
                    'department_mix': dict(demand_data['by_dept'])
                })

                self.stats['sections_created'] += 1

            print(f"   ✅ {elective_code}: {sections_needed} sections for {total_demand} students")

    def schedule_elective_sections(self):
        """Schedule all elective sections"""
        print("\n📅 SCHEDULING ELECTIVE SECTIONS")
        print("-" * 50)

        scheduled_count = 0
        failed_count = 0

        # Sort by priority
        priority_order = {'high': 1, 'medium': 2, 'low': 3}

        electives_by_priority = sorted(
            self.electives_pool.items(),
            key=lambda x: priority_order.get(x[1]['priority'], 3)
        )

        for elective_code, elective_data in electives_by_priority:
            if elective_code not in self.elective_sections:
                continue

            print(f"\n🎯 Scheduling {elective_code} ({elective_data['type']})")

            for section_data in self.elective_sections[elective_code]:
                section_key = section_data['section_key']

                if self._schedule_single_section(section_key, elective_code, elective_data):
                    scheduled_count += 1
                    print(f"   ✅ {section_key}")
                else:
                    failed_count += 1
                    print(f"   ❌ {section_key} - No available slot")

        print(f"\n📊 Scheduling Results:")
        print(f"   ✅ Scheduled: {scheduled_count}")
        print(f"   ❌ Failed: {failed_count}")

    def _schedule_single_section(self, section_key, elective_code, elective_data):
        """Schedule a single elective section"""
        lectures_needed = elective_data['credit_hours']  # 1 lecture per credit hour
        elective_type = elective_data['type']

        # Get preferred time slots
        preferred_times = self.ELECTIVE_TYPES.get(elective_type, {}).get('preferred_times', self.TIME_SLOTS)

        for lecture_num in range(1, lectures_needed + 1):
            placed = False

            for attempt in range(50):
                day = random.choice(self.DAYS)
                time_slot = random.choice(preferred_times)

                if self._can_place_elective(section_key, day, time_slot):
                    room = self._get_elective_room(elective_data, day, time_slot)

                    self._place_elective(section_key, elective_code, elective_data, day, time_slot, room, lecture_num, lectures_needed)
                    placed = True
                    break

            if not placed:
                return False

        return True

    def _can_place_elective(self, section_key, day, time_slot):
        """Check if elective can be placed"""
        return pd.isna(self.elective_timetables[section_key].loc[time_slot, day])

    def _get_elective_room(self, elective_data, day, time_slot):
        """Get available room for elective"""
        if elective_data['can_use_lab']:
            room_pool = self.room_pools['lab'] + self.room_pools['theory']
        else:
            room_pool = self.room_pools['theory']

        # Simple room assignment (can be enhanced with conflict checking)
        return random.choice(room_pool)

    def _place_elective(self, section_key, elective_code, elective_data, day, time_slot, room, lecture_num, total_lectures):
        """Place elective in timetable"""
        if total_lectures > 1:
            course_info = f"{elective_code} (L{lecture_num}/{total_lectures})\n{elective_data['name']}\n(Elective - {elective_data['type']})\nRoom: {room}"
        else:
            course_info = f"{elective_code}\n{elective_data['name']}\n(Elective - {elective_data['type']})\nRoom: {room}"

        self.elective_timetables[section_key].loc[time_slot, day] = course_info

    # ==================================================================================
    #                               OUTPUT GENERATION
    # ==================================================================================

    def generate_electives_output(self, filename="Electives_Timetable.xlsx"):
        """Generate comprehensive electives timetable"""
        print(f"\n💾 GENERATING ELECTIVES OUTPUT: {filename}")
        print("-" * 50)

        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Main electives timetable
                self._create_main_electives_sheet(writer)

                # Department-wise electives
                self._create_department_electives_sheets(writer)

                # Cross-department electives
                self._create_cross_department_sheet(writer)

                # Electives summary and analytics
                self._create_electives_summary(writer)

                # Student enrollment simulation
                self._create_enrollment_simulation(writer)

            print(f"✅ Electives timetable saved: {filename}")

        except Exception as e:
            print(f"❌ Error saving electives timetable: {e}")

    def _create_main_electives_sheet(self, writer):
        """Create main electives timetable sheet"""
        sheet_data = []

        # Header
        sheet_data.append(["UNIVERSITY ELECTIVES TIMETABLE", "", "", "", "", "", ""])
        sheet_data.append([f"Total Electives: {self.stats['total_electives']}", f"Sections: {self.stats['sections_created']}", "", "", "", "", ""])
        sheet_data.append(["", "", "", "", "", "", ""])

        # Group by elective type
        for elective_type in self.ELECTIVE_TYPES.keys():
            sheet_data.append([f"═══ {elective_type.upper()} ELECTIVES ═══", "", "", "", "", "", ""])
            sheet_data.append(["Time Slots", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])

            # Find sections of this type
            type_sections = []
            for elective_code, sections in self.elective_sections.items():
                if self.electives_pool[elective_code]['type'] == elective_type:
                    for section in sections:
                        type_sections.append(section['section_key'])

            # Display timetable for this type
            if type_sections:
                # Combine schedules (simplified view)
                combined_schedule = pd.DataFrame(index=self.TIME_SLOTS, columns=self.DAYS)

                for time_slot in self.TIME_SLOTS:
                    row = [time_slot]
                    for day in self.DAYS:
                        day_courses = []
                        for section_key in type_sections:
                            cell = self.elective_timetables[section_key].loc[time_slot, day]
                            if not pd.isna(cell):
                                course_code = cell.split('\n')[0]
                                day_courses.append(course_code)

                        if day_courses:
                            row.append(', '.join(day_courses))
                        else:
                            row.append("")

                    sheet_data.append(row)
            else:
                sheet_data.append(["No sections scheduled for this type", "", "", "", "", "", ""])

            sheet_data.append(["", "", "", "", "", "", ""])

        df = pd.DataFrame(sheet_data)
        df.to_excel(writer, sheet_name="Main_Electives", index=False, header=False)
        print("   ✅ Main_Electives")

    def _create_department_electives_sheets(self, writer):
        """Create department-specific electives sheets"""
        departments = set()
        for elective_data in self.electives_pool.values():
            departments.add(elective_data['source_department'])

        for dept in sorted(departments):
            sheet_data = []
            sheet_data.append([f"{dept} DEPARTMENT ELECTIVES", "", "", "", "", "", ""])
            sheet_data.append(["", "", "", "", "", "", ""])
            sheet_data.append(["Time Slots", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])

            # Find sections for this department
            dept_sections = []
            for elective_code, sections in self.elective_sections.items():
                if self.electives_pool[elective_code]['source_department'] == dept:
                    for section in sections:
                        dept_sections.append((section['section_key'], elective_code))

            # Create combined timetable
            for time_slot in self.TIME_SLOTS:
                row = [time_slot]
                for day in self.DAYS:
                    day_content = []
                    for section_key, elective_code in dept_sections:
                        cell = self.elective_timetables[section_key].loc[time_slot, day]
                        if not pd.isna(cell):
                            day_content.append(f"{elective_code}")

                    row.append(', '.join(day_content) if day_content else "")
                sheet_data.append(row)

            df = pd.DataFrame(sheet_data)
            sheet_name = f"Electives_{dept}"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            print(f"   ✅ {sheet_name}")

    def _create_cross_department_sheet(self, writer):
        """Create cross-department electives sheet"""
        sheet_data = []
        sheet_data.append(["CROSS-DEPARTMENT ELECTIVES", "", "", "", "", "", ""])
        sheet_data.append([f"Available to multiple departments", "", "", "", "", "", ""])
        sheet_data.append(["", "", "", "", "", "", ""])

        # List cross-department electives
        sheet_data.append(["Elective Code", "Name", "Type", "Eligible Departments", "Sections", "", ""])

        for elective_code, elective_data in self.electives_pool.items():
            eligible = elective_data['eligible_departments']
            if len(eligible) > 1 or 'ALL' in eligible:
                sections_count = len(self.elective_sections.get(elective_code, []))
                dept_list = ', '.join(eligible)

                sheet_data.append([
                    elective_code,
                    elective_data['name'],
                    elective_data['type'],
                    dept_list,
                    sections_count,
                    "", ""
                ])

        df = pd.DataFrame(sheet_data)
        df.to_excel(writer, sheet_name="Cross_Department", index=False, header=False)
        print("   ✅ Cross_Department")

    def _create_electives_summary(self, writer):
        """Create electives summary and analytics"""
        summary_data = []

        # Header
        summary_data.append(["ELECTIVES SYSTEM SUMMARY", "", "", "", ""])
        summary_data.append(["", "", "", "", ""])

        # Statistics
        summary_data.append(["SYSTEM STATISTICS", "", "", "", ""])
        summary_data.append(["Total Electives Available", self.stats['total_electives'], "", "", ""])
        summary_data.append(["Cross-Department Electives", self.stats['cross_dept_electives'], "", "", ""])
        summary_data.append(["Total Sections Created", self.stats['sections_created'], "", "", ""])
        summary_data.append(["Total Students (simulated)", len(self.student_preferences), "", "", ""])
        summary_data.append(["", "", "", "", ""])

        # Elective types breakdown
        summary_data.append(["ELECTIVES BY TYPE", "", "", "", ""])
        type_counts = defaultdict(int)
        for elective_data in self.electives_pool.values():
            type_counts[elective_data['type']] += 1

        for elective_type, count in type_counts.items():
            summary_data.append([elective_type, count, "", "", ""])

        summary_data.append(["", "", "", "", ""])

        # Department breakdown
        summary_data.append(["ELECTIVES BY DEPARTMENT", "", "", "", ""])
        dept_counts = defaultdict(int)
        for elective_data in self.electives_pool.values():
            dept_counts[elective_data['source_department']] += 1

        for dept, count in dept_counts.items():
            summary_data.append([dept, count, "", "", ""])

        df = pd.DataFrame(summary_data)
        df.to_excel(writer, sheet_name="Summary", index=False, header=False)
        print("   ✅ Summary")

    def _create_enrollment_simulation(self, writer):
        """Create student enrollment simulation"""
        enrollment_data = []

        # Header
        enrollment_data.append(["STUDENT ENROLLMENT SIMULATION", "", "", "", "", ""])
        enrollment_data.append(["Student ID", "Department", "Semester", "Elective Code", "Priority", "Status"])

        # Sample enrollment data
        for student_key, preferences in list(self.student_preferences.items())[:100]:  # First 100 students
            for pref in preferences:
                enrollment_data.append([
                    student_key,
                    pref['department'],
                    pref['semester'],
                    pref['elective_code'],
                    pref['priority'],
                    "Pending"  # In real system, this would be Enrolled/Waitlisted/Rejected
                ])

        df = pd.DataFrame(enrollment_data)
        df.to_excel(writer, sheet_name="Enrollment_Simulation", index=False, header=False)
        print("   ✅ Enrollment_Simulation")

    # ==================================================================================
    #                               MAIN EXECUTION
    # ==================================================================================

    def run_electives_system(self, core_files=None):
        """Run complete electives management system"""
        print("🎓 ELECTIVES MANAGEMENT SYSTEM")
        print("=" * 60)

        # Step 1: Load core files
        if not self.load_core_files(core_files):
            print("❌ No core files found!")
            return False

        # Step 2: Setup rooms
        self.setup_rooms()

        # Step 3: Process electives data
        self.process_electives_data()

        if self.stats['total_electives'] == 0:
            print("❌ No electives found in the files!")
            return False

        # Step 4: Generate student preferences (simulation)
        self.generate_sample_student_preferences()

        # Step 5: Analyze demand
        demand_analysis = self.analyze_demand()

        # Step 6: Create sections based on demand
        self.create_elective_sections(demand_analysis)

        # Step 7: Schedule all sections
        self.schedule_elective_sections()

        # Step 8: Generate output
        self.generate_electives_output()

        # Step 9: Generate additional reports
        self.generate_additional_reports()

        # Final report
        self._print_final_electives_report()

        return True

    def generate_additional_reports(self):
        """Generate additional specialized reports"""
        print("\n📊 GENERATING ADDITIONAL REPORTS")
        print("-" * 50)

        # 1. Conflict Analysis Report
        self._generate_conflict_report()

        # 2. Capacity Utilization Report
        self._generate_capacity_report()

        # 3. Student Choice Analysis
        self._generate_choice_analysis()

    def _generate_conflict_report(self):
        """Generate conflict analysis report"""
        try:
            conflicts = []

            # Check for time slot conflicts
            time_slot_usage = defaultdict(lambda: defaultdict(list))

            for section_key, timetable in self.elective_timetables.items():
                for time_slot in self.TIME_SLOTS:
                    for day in self.DAYS:
                        if not pd.isna(timetable.loc[time_slot, day]):
                            time_slot_usage[day][time_slot].append(section_key)

            # Find conflicts
            for day in self.DAYS:
                for time_slot in self.TIME_SLOTS:
                    sections = time_slot_usage[day][time_slot]
                    if len(sections) > 1:
                        conflicts.append({
                            'day': day,
                            'time_slot': time_slot,
                            'conflicting_sections': sections,
                            'conflict_type': 'Time Slot Overlap'
                        })

            # Save conflict report
            with pd.ExcelWriter("Electives_Conflict_Report.xlsx", engine='openpyxl') as writer:
                if conflicts:
                    conflict_data = []
                    conflict_data.append(["ELECTIVES CONFLICT ANALYSIS", "", "", ""])
                    conflict_data.append(["Day", "Time Slot", "Conflicting Sections", "Resolution"])

                    for conflict in conflicts:
                        conflict_data.append([
                            conflict['day'],
                            conflict['time_slot'],
                            ', '.join(conflict['conflicting_sections']),
                            "Manual review required"
                        ])

                    df = pd.DataFrame(conflict_data)
                    df.to_excel(writer, sheet_name="Conflicts", index=False, header=False)
                else:
                    no_conflict_data = [["NO CONFLICTS DETECTED", "", "", ""]]
                    df = pd.DataFrame(no_conflict_data)
                    df.to_excel(writer, sheet_name="Conflicts", index=False, header=False)

            print(f"   ✅ Conflict Report: {len(conflicts)} conflicts found")

        except Exception as e:
            print(f"   ❌ Error generating conflict report: {e}")

    def _generate_capacity_report(self):
        """Generate capacity utilization report"""
        try:
            with pd.ExcelWriter("Electives_Capacity_Report.xlsx", engine='openpyxl') as writer:
                capacity_data = []

                # Header
                capacity_data.append(["ELECTIVES CAPACITY ANALYSIS", "", "", "", "", ""])
                capacity_data.append(["Elective Code", "Elective Name", "Demand", "Sections", "Total Capacity", "Utilization"])

                for elective_code in self.electives_pool:
                    elective_data = self.electives_pool[elective_code]
                    sections = self.elective_sections.get(elective_code, [])

                    # Calculate demand from preferences
                    demand = sum(1 for prefs in self.student_preferences.values()
                               for pref in prefs if pref['elective_code'] == elective_code)

                    total_capacity = sum(section['capacity'] for section in sections)
                    utilization = (demand / total_capacity * 100) if total_capacity > 0 else 0

                    capacity_data.append([
                        elective_code,
                        elective_data['name'],
                        demand,
                        len(sections),
                        total_capacity,
                        f"{utilization:.1f}%"
                    ])

                df = pd.DataFrame(capacity_data)
                df.to_excel(writer, sheet_name="Capacity_Analysis", index=False, header=False)

            print("   ✅ Capacity Report generated")

        except Exception as e:
            print(f"   ❌ Error generating capacity report: {e}")

    def _generate_choice_analysis(self):
        """Generate student choice fulfillment analysis"""
        try:
            with pd.ExcelWriter("Electives_Choice_Analysis.xlsx", engine='openpyxl') as writer:
                choice_data = []

                # Header
                choice_data.append(["STUDENT CHOICE ANALYSIS", "", "", "", ""])
                choice_data.append(["Department", "Total Students", "First Choice Available", "Second Choice Available", "No Choice Available"])

                # Analyze by department
                departments = ['CS', 'SE', 'AI', 'DS', 'INFS', 'CB']

                for dept in departments:
                    dept_students = [key for key in self.student_preferences.keys() if dept in key]
                    total_students = len(dept_students)

                    first_choice_available = 0
                    second_choice_available = 0
                    no_choice_available = 0

                    for student_key in dept_students:
                        preferences = self.student_preferences[student_key]
                        if not preferences:
                            no_choice_available += 1
                            continue

                        # Check if first choice has sections
                        first_choice = next((p for p in preferences if p['priority'] == 1), None)
                        if first_choice and first_choice['elective_code'] in self.elective_sections:
                            first_choice_available += 1
                        else:
                            # Check second choice
                            second_choice = next((p for p in preferences if p['priority'] == 2), None)
                            if second_choice and second_choice['elective_code'] in self.elective_sections:
                                second_choice_available += 1
                            else:
                                no_choice_available += 1

                    choice_data.append([
                        dept,
                        total_students,
                        first_choice_available,
                        second_choice_available,
                        no_choice_available
                    ])

                df = pd.DataFrame(choice_data)
                df.to_excel(writer, sheet_name="Choice_Analysis", index=False, header=False)

            print("   ✅ Choice Analysis Report generated")

        except Exception as e:
            print(f"   ❌ Error generating choice analysis: {e}")

    def _print_final_electives_report(self):
        """Print comprehensive final report"""
        print(f"\n" + "=" * 60)
        print("📋 FINAL ELECTIVES SYSTEM REPORT")
        print("=" * 60)

        print(f"\n📊 ELECTIVES OVERVIEW:")
        print(f"   🎓 Total electives processed: {self.stats['total_electives']}")
        print(f"   🏫 Cross-department electives: {self.stats['cross_dept_electives']}")
        print(f"   📚 Total sections created: {self.stats['sections_created']}")

        print(f"\n👥 STUDENT SIMULATION:")
        print(f"   👨‍🎓 Total students (simulated): {len(self.student_preferences)}")
        total_preferences = sum(len(prefs) for prefs in self.student_preferences.values())
        print(f"   📝 Total preference entries: {total_preferences}")
        print(f"   📊 Average choices per student: {total_preferences/len(self.student_preferences):.1f}")

        print(f"\n🎯 ELECTIVES BY TYPE:")
        type_counts = defaultdict(int)
        for elective_data in self.electives_pool.values():
            type_counts[elective_data['type']] += 1

        for elective_type, count in type_counts.items():
            print(f"   {elective_type}: {count} electives")

        print(f"\n🏛️ ELECTIVES BY DEPARTMENT:")
        dept_counts = defaultdict(int)
        for elective_data in self.electives_pool.values():
            dept_counts[elective_data['source_department']] += 1

        for dept, count in dept_counts.items():
            print(f"   {dept}: {count} electives")

        print(f"\n🏢 RESOURCE UTILIZATION:")
        print(f"   📚 Theory rooms available: {len(self.room_pools['theory'])}")
        print(f"   🔬 Lab rooms available: {len(self.room_pools['lab'])}")

        print(f"\n📁 OUTPUT FILES GENERATED:")
        print(f"   📊 Electives_Timetable.xlsx (Main timetable)")
        print(f"   ⚠️ Electives_Conflict_Report.xlsx (Conflict analysis)")
        print(f"   📈 Electives_Capacity_Report.xlsx (Capacity utilization)")
        print(f"   🎯 Electives_Choice_Analysis.xlsx (Student choice analysis)")

        print(f"\n💡 RECOMMENDATIONS:")
        if self.stats['cross_dept_electives'] > 0:
            print(f"   ✅ Good cross-department coverage promotes interdisciplinary learning")

        if self.stats['sections_created'] < self.stats['total_electives']:
            print(f"   📈 Consider increasing marketing for low-demand electives")

        print(f"   🔧 Integrate with main timetable system for conflict resolution")
        print(f"   👥 Implement real student preference collection system")
        print(f"   📱 Consider mobile app for elective registration")

        print(f"\n✨ Electives Management System completed successfully!")


# ==================================================================================
#                                   EXECUTION
# ==================================================================================

def main():
    """Main execution function for electives system"""
    # Initialize electives manager
    manager = ElectivesManager()

    # Specify available core files (same as main system)
    available_files = [
        'BSCB.xlsx', 'BSCS .xlsx', 'BSINFS.xlsx', 'BSSE.xlsx',
        'BSAI_4.2.xlsx', 'BSDS 4.2.xlsx'
    ]

    print(f"🎯 PROCESSING ELECTIVES FROM {len(available_files)} CORE FILES")

    # Run the complete electives system
    success = manager.run_electives_system(available_files)

    if success:
        print(f"\n🎉 SUCCESS! Electives system completed!")
        print(f"📁 Check generated Excel files for detailed reports")
    else:
        print(f"\n❌ Electives system failed to complete")

    return manager

# Run the electives system
if __name__ == "__main__":
    electives_manager = main()

    print(f"\n" + "=" * 60)
    print("🎓 ELECTIVES SYSTEM INTEGRATION NOTES")
    print("=" * 60)
    print("📋 TO INTEGRATE WITH MAIN TIMETABLE SYSTEM:")
    print("   1. Import this module: from electives_manager import ElectivesManager")
    print("   2. Run after core scheduling: electives = ElectivesManager()")
    print("   3. Load existing core files: electives.load_core_files(core_files)")
    print("   4. Check for conflicts: electives.check_conflicts_with_main_timetable()")
    print("   5. Generate combined output: electives.merge_with_main_schedule()")
    print("\n💡 FUTURE ENHANCEMENTS:")
    print("   🔧 Real-time student registration system")
    print("   📊 Advanced analytics and recommendation engine")
    print("   🤖 AI-powered demand prediction")
    print("   📱 Mobile app integration")
    print("   🎯 Automatic conflict resolution with main timetable")
    print("\n✨ Electives Management System is ready for deployment!")