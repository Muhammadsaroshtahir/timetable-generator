import pandas as pd
import numpy as np
from collections import defaultdict
import random
import math
import os
from pathlib import Path

# ==================================================================================
#                      ULTIMATE 12-FILE TIMETABLE GENERATOR
# ==================================================================================

class TimetableGenerator:
    """
    Ultimate Timetable Generator for 12 files (6 core + 6 cohort)
    Supports scalable daily classes (3-4), unified time system, and smart optimization
    """

    def __init__(self):
        # ---------------------- TIME CONFIGURATION ----------------------
        self.TIME_SLOTS = [
            '08:00-09:15', '09:30-10:45', '11:00-12:15',
            '12:30-01:45', '02:00-03:15', '03:30-04:45', '05:00-06:15'
        ]

        self.LAB_COMBINATIONS = [
            ['08:00-09:15', '09:30-10:45'], ['09:30-10:45', '11:00-12:15'],
            ['11:00-12:15', '12:30-01:45'], ['12:30-01:45', '02:00-03:15'],
            ['02:00-03:15', '03:30-04:45'], ['03:30-04:45', '05:00-06:15']
        ]

        self.DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

        # ---------------------- CONSTRAINTS ----------------------
        self.MIN_CLASSES_PER_DAY = 3
        self.MAX_CLASSES_PER_DAY = 4
        self.STUDENTS_PER_SECTION = 50

        # ---------------------- DATA STORAGE ----------------------
        self.core_files = {}
        self.cohort_files = {}
        self.rooms = {'theory': [], 'lab': [], 'special': {}}
        self.timetables = defaultdict(lambda: pd.DataFrame(index=self.TIME_SLOTS, columns=self.DAYS))
        self.daily_counts = defaultdict(lambda: defaultdict(int))
        self.lecture_counts = defaultdict(lambda: defaultdict(int))
        self.room_bookings = defaultdict(lambda: defaultdict(set))
        self.batch_info = {}
        self.stats = {'placed': 0, 'failed': 0, 'cohort': 0}

        print("🚀 Ultimate 12-File Timetable Generator Initialized")
        print(f"📅 Time Slots: {len(self.TIME_SLOTS)} | Lab Options: {len(self.LAB_COMBINATIONS)}")
        print(f"⚖️ Daily Classes: {self.MIN_CLASSES_PER_DAY}-{self.MAX_CLASSES_PER_DAY}")

    # ==================================================================================
    #                               FILE OPERATIONS
    # ==================================================================================

    def load_files(self, file_list=None):
        """Load all Excel files from directory or specified list"""
        print("\n📂 LOADING FILES")
        print("-" * 40)

        if file_list is None:
            file_list = list(Path('.').glob('*.xlsx'))
            file_list = [f for f in file_list if not f.name.startswith('~')]
        else:
            # Filter to only existing files
            existing_files = []
            for file_path in file_list:
                if os.path.exists(file_path):
                    existing_files.append(file_path)
                else:
                    print(f"⚠️ File not found: {file_path}")
            file_list = existing_files

        print(f"📁 Processing {len(file_list)} Excel files")

        for file_path in file_list:
            self._load_single_file(file_path)

        print(f"\n✅ Successfully loaded:")
        print(f"   📚 Core files: {len(self.core_files)}")
        print(f"   📅 Cohort files: {len(self.cohort_files)}")
        print(f"   🎯 Total: {len(self.core_files) + len(self.cohort_files)} files")

        return len(self.core_files) > 0

    def _load_single_file(self, file_path):
        """Load individual Excel file with error handling"""
        try:
            file_name = Path(file_path).name
            print(f"📄 Loading {file_name}...")

            # Determine file type
            is_cohort = any(word in file_name.lower() for word in ['cohort', 'fixed', 'schedule'])

            # Load Excel file
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

            # Store file data
            file_data = {
                'name': file_name,
                'sheets': sheets,
                'department': self._extract_department(file_name)
            }

            if is_cohort:
                self.cohort_files[file_name] = file_data
                print(f"   📅 Cohort file: {len(sheets)} sheets")
            else:
                self.core_files[file_name] = file_data
                print(f"   📚 Core file: {len(sheets)} sheets")

        except Exception as e:
            print(f"   ❌ Failed to load {file_path}: {e}")

    def _extract_department(self, filename):
        """Extract department name from filename"""
        name = Path(filename).stem.upper()

        if 'COHORT' in name:
            # Extract from cohort_DEPT_X.X.xlsx
            parts = name.split('_')
            if len(parts) >= 2:
                return parts[1]  # CB, CS, INFS, SE, DS, AI
            return 'UNKNOWN'
        else:
            # Extract from core files
            # Handle variations like BSCS .xlsx, BSAI_4.2.xlsx, BSDS 4.2.xlsx
            name = name.replace(' ', '').replace('_4.2', '').replace('.', '')

            # Map common department codes
            dept_mapping = {
                'BSCB': 'CB',
                'BSCS': 'CS',
                'BSINFS': 'INFS',
                'BSSE': 'SE',
                'BSAI42': 'AI',
                'BSAI': 'AI',
                'BSDS42': 'DS',
                'BSDS': 'DS'
            }

            return dept_mapping.get(name, name)

    # ==================================================================================
    #                               SETUP OPERATIONS
    # ==================================================================================

    def setup_rooms(self):
        """Setup room pools from all core files"""
        print("\n🏢 SETTING UP ROOMS")
        print("-" * 40)

        theory_rooms = set()
        lab_rooms = set()
        special_rooms = {}

        for file_data in self.core_files.values():
            # Load standard rooms
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

            # Load special labs
            if 'SpecialLabs' in file_data['sheets']:
                special_df = file_data['sheets']['SpecialLabs']
                for _, row in special_df.iterrows():
                    course = str(row.get('course_code', '')).strip()
                    rooms_str = str(row.get('lab_rooms', '')).strip()
                    if course and rooms_str:
                        special_rooms[course] = [r.strip() for r in rooms_str.split(',')]

        # Set defaults if no rooms found
        self.rooms['theory'] = list(theory_rooms) or [f'T{i:02d}' for i in range(1, 21)]
        self.rooms['lab'] = list(lab_rooms) or [f'L{i:02d}' for i in range(1, 16)]
        self.rooms['special'] = special_rooms

        print(f"📚 Theory rooms: {len(self.rooms['theory'])}")
        print(f"🔬 Lab rooms: {len(self.rooms['lab'])}")
        print(f"🎯 Special labs: {len(self.rooms['special'])}")

    def analyze_capacity(self):
        """Analyze student capacity across all departments"""
        print("\n👥 ANALYZING STUDENT CAPACITY")
        print("-" * 40)

        for file_name, file_data in self.core_files.items():
            if 'StudentCapacity' not in file_data['sheets']:
                continue

            dept = file_data['department']
            capacity_df = file_data['sheets']['StudentCapacity']

            print(f"\n📊 {dept} Department:")

            for _, row in capacity_df.iterrows():
                try:
                    semester = int(row.get('semester', 0))
                    student_count = self._safe_int(row.get('student_count', 0))

                    if semester and student_count > 0:
                        sections = math.ceil(student_count / self.STUDENTS_PER_SECTION)
                        section_names = [f"{dept}_S{semester}_Sec{chr(65+i)}" for i in range(sections)]

                        batch_key = f"{dept}_Sem{semester}"
                        self.batch_info[batch_key] = {
                            'department': dept,
                            'semester': semester,
                            'students': student_count,
                            'sections': sections,
                            'section_names': section_names
                        }

                        print(f"   Sem {semester}: {student_count} students → {sections} sections")

                except Exception as e:
                    print(f"   ⚠️ Error processing capacity: {e}")

    def _safe_int(self, value):
        """Safely convert value to integer"""
        if pd.isna(value) or str(value).lower() in ['nil', 'null', '', 'nan']:
            return 0
        try:
            return int(float(value))
        except:
            return 0

    # ==================================================================================
    #                               SCHEDULING OPERATIONS
    # ==================================================================================

    def schedule_cohort_courses(self):
        """Schedule fixed cohort courses with enhanced processing"""
        print("\n📅 SCHEDULING COHORT COURSES")
        print("-" * 40)

        total_cohort_placed = 0

        for file_name, file_data in self.cohort_files.items():
            dept = file_data['department']
            main_sheet = self._get_main_sheet(file_data['sheets'])

            if not main_sheet:
                print(f"   ⚠️ No main sheet found in {file_name}")
                continue

            print(f"\n📋 Processing {dept} cohort from {file_name}:")
            cohort_df = file_data['sheets'][main_sheet]

            if cohort_df.empty:
                print(f"   ⚠️ Empty cohort data in {file_name}")
                continue

            print(f"   📊 Found {len(cohort_df)} rows of cohort data")

            # Debug: Show first few rows
            print(f"   🔍 Column names: {list(cohort_df.columns)}")
            print(f"   🔍 Sample data:")
            for i in range(min(3, len(cohort_df))):
                print(f"      Row {i}: {dict(cohort_df.iloc[i])}")

            courses_in_file = 0

            for idx, row in cohort_df.iterrows():
                try:
                    # More flexible column name matching
                    semester_col = None
                    for col in ['CohortSemester', 'Semester', 'semester', 'Sem']:
                        if col in row and pd.notna(row[col]):
                            semester_col = col
                            break

                    if not semester_col:
                        continue

                    semester = self._safe_int(row[semester_col])
                    if semester == 0:
                        continue

                    # Flexible course code matching
                    course_code_col = None
                    for col in ['CourseCode', 'Course Code', 'C-CODE', 'course_code']:
                        if col in row and pd.notna(row[col]):
                            course_code_col = col
                            break

                    if not course_code_col:
                        continue

                    course_code = str(row[course_code_col]).strip()

                    # Flexible course name matching
                    course_name_col = None
                    for col in ['CourseName', 'Course Name', 'course_name', 'CourseTitle']:
                        if col in row and pd.notna(row[col]):
                            course_name_col = col
                            break

                    course_name = str(row[course_name_col]).strip() if course_name_col else course_code

                    # Flexible section matching
                    section_col = None
                    for col in ['Section', 'section', 'Sec', 'SEC']:
                        if col in row and pd.notna(row[col]):
                            section_col = col
                            break

                    section = str(row[section_col]).strip() if section_col else f"Default_{idx}"

                    # Flexible capacity matching
                    capacity_col = None
                    for col in ['Capacity', 'capacity', 'Cap', 'Students']:
                        if col in row and pd.notna(row[col]):
                            capacity_col = col
                            break

                    capacity = self._safe_int(row[capacity_col]) if capacity_col else 50
                    if capacity == 0:
                        capacity = 50

                    if not course_code or course_code == 'nan':
                        continue

                    section_key = f"Cohort_{dept}_S{semester}_{section}"

                    # Initialize timetable if needed
                    if section_key not in self.timetables:
                        print(f"   🆕 Creating new section: {section_key}")

                    # Process weekly schedule - be more flexible with day names
                    day_columns = []
                    for day in self.DAYS:
                        # Check multiple possible day column names
                        day_variations = [day, day.lower(), day[:3], day[:3].lower()]
                        for variation in day_variations:
                            if variation in row.index:
                                day_columns.append((day, variation))
                                break

                    course_placed = False
                    for day, day_col in day_columns:
                        time_slot = str(row[day_col]).strip() if pd.notna(row[day_col]) else ""

                        if time_slot and time_slot != 'nan' and time_slot != '':
                            # Try to match time slot format
                            if time_slot in self.TIME_SLOTS:
                                if pd.isna(self.timetables[section_key].loc[time_slot, day]):
                                    course_info = f"{course_code}\n{course_name}\n(Fixed)\nCap: {capacity}"
                                    self.timetables[section_key].loc[time_slot, day] = course_info
                                    self.daily_counts[section_key][day] += 1
                                    total_cohort_placed += 1
                                    course_placed = True
                                    print(f"   ✅ {course_code} → {section_key} → {day} {time_slot}")
                                else:
                                    print(f"   ⚠️ Conflict: {course_code} → {section_key} → {day} {time_slot}")
                            else:
                                print(f"   ⚠️ Invalid time slot '{time_slot}' for {course_code}")

                    if course_placed:
                        courses_in_file += 1

                except Exception as e:
                    print(f"   ⚠️ Error processing cohort row {idx}: {e}")

            print(f"   📊 Placed {courses_in_file} courses from {file_name}")

        self.stats['cohort'] = total_cohort_placed
        print(f"\n📈 Total cohort courses placed: {total_cohort_placed}")

        if total_cohort_placed == 0:
            print("❌ No cohort courses were placed - check data format and column names")

    def schedule_core_courses(self):
        """Schedule core courses with optimization"""
        print("\n📚 SCHEDULING CORE COURSES")
        print("-" * 40)

        for file_name, file_data in self.core_files.items():
            if 'Roadmap' not in file_data['sheets']:
                continue

            dept = file_data['department']
            courses_df = file_data['sheets']['Roadmap']

            print(f"\n🎯 {dept} Department ({len(courses_df)} courses):")

            for _, row in courses_df.iterrows():
                self._schedule_single_course(row, dept)

    def _schedule_single_course(self, course_row, department):
        """Schedule a single course across all its sections"""
        try:
            semester = int(course_row.get('semester', 0))
            course_code = str(course_row.get('course_code', '')).strip()
            course_name = str(course_row.get('course_name', '')).strip()
            is_lab = bool(course_row.get('is_lab', False))
            lectures_needed = int(course_row.get('times_needed', 1))

            if not course_code or course_code == 'nan':
                return

            batch_key = f"{department}_Sem{semester}"
            if batch_key not in self.batch_info:
                return

            # Schedule for each section
            for section_name in self.batch_info[batch_key]['section_names']:
                self._place_course_lectures(section_name, course_code, course_name,
                                          is_lab, lectures_needed)

        except Exception as e:
            print(f"   ⚠️ Error scheduling {course_code}: {e}")

    def _place_course_lectures(self, section_key, course_code, course_name, is_lab, lectures_needed):
        """Place all lectures for a course in a section"""
        for lecture_num in range(1, lectures_needed + 1):
            placed = False

            # Choose time slots
            time_options = self.LAB_COMBINATIONS if is_lab else [[slot] for slot in self.TIME_SLOTS]

            # Try placement
            for attempt in range(100):
                day = random.choice(self.DAYS)
                time_slots = random.choice(time_options)

                if self._can_place_course(section_key, course_code, day, time_slots, lectures_needed):
                    room = self._get_room(course_code, is_lab, day, time_slots[0])
                    self._place_course(section_key, course_code, course_name, day,
                                     time_slots, room, lecture_num, lectures_needed)
                    self.stats['placed'] += 1
                    placed = True
                    break

            if not placed:
                self.stats['failed'] += 1

    def _can_place_course(self, section_key, course_code, day, time_slots, max_lectures):
        """Check if course can be placed"""
        # Check lecture limit
        if self.lecture_counts[section_key][course_code] >= max_lectures:
            return False

        # Check daily limit (scalable 3-4)
        if self.daily_counts[section_key][day] >= self.MAX_CLASSES_PER_DAY:
            return False

        # Check time slot availability
        for slot in time_slots:
            if not pd.isna(self.timetables[section_key].loc[slot, day]):
                return False

        return True

    def _place_course(self, section_key, course_code, course_name, day, time_slots, room, lecture_num, total):
        """Place course in timetable"""
        # Create course info
        if total > 1:
            info = f"{course_code} (L{lecture_num}/{total})\n{course_name}\nRoom: {room}"
        else:
            info = f"{course_code}\n{course_name}\nRoom: {room}"

        # Add lab timing
        if len(time_slots) > 1:
            start = time_slots[0].split('-')[0]
            end = time_slots[-1].split('-')[1]
            info += f"\n(Lab: {start}-{end})"

        # Place in timetable
        for slot in time_slots:
            self.timetables[section_key].loc[slot, day] = info

        # Update counters
        self.daily_counts[section_key][day] += 1
        self.lecture_counts[section_key][course_code] += 1

    def _get_room(self, course_code, is_lab, day, time_slot):
        """Get available room with conflict management"""
        # Determine room pool
        if is_lab and course_code in self.rooms['special']:
            pool = self.rooms['special'][course_code]
        elif is_lab:
            pool = self.rooms['lab']
        else:
            pool = self.rooms['theory']

        # Find available room
        booked = self.room_bookings[day][time_slot]
        available = [room for room in pool if room not in booked]

        if available:
            room = random.choice(available)
            self.room_bookings[day][time_slot].add(room)
            return room
        else:
            return f"{random.choice(pool)}*"  # Mark conflict

    def _get_main_sheet(self, sheets):
        """Get main sheet from cohort file"""
        if 'Sheet1' in sheets:
            return 'Sheet1'
        elif sheets:
            return list(sheets.keys())[0]
        return None

    # ==================================================================================
    #                               OUTPUT OPERATIONS
    # ==================================================================================

    def generate_output(self, filename="Ultimate_12File_Timetable.xlsx"):
        """Generate comprehensive Excel output"""
        print(f"\n💾 GENERATING OUTPUT: {filename}")
        print("-" * 40)

        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Department sheets
                for batch_key, batch_data in self.batch_info.items():
                    self._create_department_sheet(writer, batch_key, batch_data)

                # Cohort sheet
                self._create_cohort_sheet(writer)

                # Summary sheet
                self._create_summary_sheet(writer)

            print(f"✅ Timetable saved: {filename}")

        except Exception as e:
            print(f"❌ Error saving: {e}")

    def _create_department_sheet(self, writer, batch_key, batch_data):
        """Create sheet for department/semester"""
        sheet_data = []

        # Header
        dept = batch_data['department']
        sem = batch_data['semester']
        sheet_data.append([f"{dept} - SEMESTER {sem}", "", "", "", "", "", ""])
        sheet_data.append([f"Students: {batch_data['students']}", f"Sections: {batch_data['sections']}", "", "", "", "", ""])
        sheet_data.append(["", "", "", "", "", "", ""])
        sheet_data.append(["Time Slots", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])

        # Sections
        for section_name in batch_data['section_names']:
            sheet_data.append([f"═══ {section_name} ═══", "", "", "", "", "", ""])

            for time_slot in self.TIME_SLOTS:
                row = [time_slot]
                for day in self.DAYS:
                    cell = self.timetables[section_name].loc[time_slot, day]
                    row.append("" if pd.isna(cell) else str(cell))
                sheet_data.append(row)

            sheet_data.append(["", "", "", "", "", "", ""])

        # Save
        df = pd.DataFrame(sheet_data)
        sheet_name = batch_key.replace('/', '_')[:31]
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        print(f"   ✅ {sheet_name}")

    def _create_cohort_sheet(self, writer):
        """Create cohort schedules sheet with enhanced debugging"""
        cohort_sections = [key for key in self.timetables.keys() if key.startswith('Cohort_')]

        print(f"   🔍 Creating cohort sheet with {len(cohort_sections)} sections")

        if not cohort_sections:
            print("   ⚠️ No cohort sections found - creating empty cohort sheet")
            sheet_data = []
            sheet_data.append(["COHORT COURSES - NO DATA FOUND", "", "", "", "", "", ""])
            sheet_data.append(["Check cohort file format and column names", "", "", "", "", "", ""])
            sheet_data.append(["Expected columns: CohortSemester, CourseCode, CourseName, Section, Capacity", "", "", "", "", "", ""])
            sheet_data.append(["Day columns: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday", "", "", "", "", "", ""])

            df = pd.DataFrame(sheet_data)
            df.to_excel(writer, sheet_name="Cohort_Schedules", index=False, header=False)
            print(f"   ✅ Empty Cohort_Schedules (diagnostic)")
            return

        sheet_data = []
        sheet_data.append(["COHORT COURSES - FIXED SCHEDULES", "", "", "", "", "", ""])
        sheet_data.append([f"Found {len(cohort_sections)} cohort sections", "", "", "", "", "", ""])
        sheet_data.append(["", "", "", "", "", "", ""])

        for section_key in sorted(cohort_sections):
            # Count actual courses in this section
            course_count = 0
            for time_slot in self.TIME_SLOTS:
                for day in self.DAYS:
                    if not pd.isna(self.timetables[section_key].loc[time_slot, day]):
                        course_count += 1

            sheet_data.append([f"═══ {section_key} ({course_count} courses) ═══", "", "", "", "", "", ""])
            sheet_data.append(["Time Slots", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])

            for time_slot in self.TIME_SLOTS:
                row = [time_slot]
                for day in self.DAYS:
                    cell = self.timetables[section_key].loc[time_slot, day]
                    row.append("" if pd.isna(cell) else str(cell))
                sheet_data.append(row)

            sheet_data.append(["", "", "", "", "", "", ""])

        df = pd.DataFrame(sheet_data)
        df.to_excel(writer, sheet_name="Cohort_Schedules", index=False, header=False)
        print(f"   ✅ Cohort_Schedules ({len(cohort_sections)} sections)")

    def _create_summary_sheet(self, writer):
        """Create comprehensive summary"""
        summary = []

        # Header
        summary.append(["ULTIMATE 12-FILE TIMETABLE SYSTEM", "", "", "", ""])
        summary.append([f"Generated for {len(self.core_files)} core + {len(self.cohort_files)} cohort files", "", "", "", ""])
        summary.append(["", "", "", "", ""])

        # Configuration
        summary.append(["SYSTEM CONFIGURATION", "", "", "", ""])
        summary.append(["Time Slots", len(self.TIME_SLOTS), "", "", ""])
        summary.append(["Daily Classes Range", f"{self.MIN_CLASSES_PER_DAY}-{self.MAX_CLASSES_PER_DAY}", "", "", ""])
        summary.append(["Students per Section", self.STUDENTS_PER_SECTION, "", "", ""])
        summary.append(["", "", "", "", ""])

        # Statistics
        summary.append(["SCHEDULING RESULTS", "", "", "", ""])
        summary.append(["Core Courses Placed", self.stats['placed'], "", "", ""])
        summary.append(["Failed Placements", self.stats['failed'], "", "", ""])
        summary.append(["Cohort Courses Placed", self.stats['cohort'], "", "", ""])

        total = self.stats['placed'] + self.stats['failed']
        success_rate = (self.stats['placed'] / total * 100) if total > 0 else 0
        summary.append(["Success Rate", f"{success_rate:.1f}%", "", "", ""])
        summary.append(["", "", "", "", ""])

        # Resources
        summary.append(["RESOURCES", "", "", "", ""])
        summary.append(["Theory Rooms", len(self.rooms['theory']), "", "", ""])
        summary.append(["Lab Rooms", len(self.rooms['lab']), "", "", ""])
        summary.append(["Special Lab Courses", len(self.rooms['special']), "", "", ""])

        df = pd.DataFrame(summary)
        df.to_excel(writer, sheet_name="Summary", index=False, header=False)
        print(f"   ✅ Summary")

    # ==================================================================================
    #                               MAIN EXECUTION
    # ==================================================================================

    def run(self, file_list=None):
        """Run complete timetable generation"""
        print("🎯 ULTIMATE 12-FILE TIMETABLE GENERATOR")
        print("=" * 60)

        # Load files
        if not self.load_files(file_list):
            print("❌ No core files found!")
            return False

        # Setup system
        self.setup_rooms()
        self.analyze_capacity()

        # Schedule courses
        self.schedule_cohort_courses()
        self.schedule_core_courses()

        # Generate output
        self.generate_output()

        # Print results
        self._print_final_report()

        return True

    def _print_final_report(self):
        """Print comprehensive final report"""
        print(f"\n" + "=" * 60)
        print("📋 FINAL SYSTEM REPORT")
        print("=" * 60)

        print(f"\n📂 FILES PROCESSED:")
        print(f"   📚 Core files: {len(self.core_files)}/6")
        print(f"   📅 Cohort files: {len(self.cohort_files)}/6")
        print(f"   🎯 Total: {len(self.core_files) + len(self.cohort_files)}/12")

        print(f"\n📊 SCHEDULING RESULTS:")
        print(f"   ✅ Core courses placed: {self.stats['placed']}")
        print(f"   ❌ Failed placements: {self.stats['failed']}")
        print(f"   📅 Cohort courses: {self.stats['cohort']}")

        total_students = sum(batch['students'] for batch in self.batch_info.values())
        total_sections = sum(batch['sections'] for batch in self.batch_info.values())

        print(f"\n🏫 CAPACITY:")
        print(f"   👥 Total students: {total_students}")
        print(f"   📊 Total sections: {total_sections}")
        print(f"   🏛️ Departments: {len(set(batch['department'] for batch in self.batch_info.values()))}")

        print(f"\n✨ 12-File system ready for expansion!")


# ==================================================================================
#                                   EXECUTION
# ==================================================================================

def main():
    """Main execution function"""
    # Initialize generator
    generator = TimetableGenerator()

    # Detect all available files from your directory
    available_files = [
        'BSCB.xlsx', 'BSCS .xlsx', 'BSINFS.xlsx', 'BSSE.xlsx',
        'BSAI_4.2.xlsx', 'BSDS 4.2.xlsx',  # Core files
        'cohort_CB_4.2.xlsx', 'cohort_CS_4.2.xlsx', 'cohort_INFS_4.2.xlsx',
        'cohort_SE_4.2.xlsx', 'cohort_DS_4.2.xlsx', 'cohort_AI_4.2.xlsx'  # Cohort files
    ]

    print(f"🎯 PROCESSING ALL AVAILABLE FILES ({len(available_files)} files)")
    print("📂 Detected files:")

    # Show available files
    core_files = [f for f in available_files if not f.startswith('cohort_')]
    cohort_files = [f for f in available_files if f.startswith('cohort_')]

    print(f"   📚 Core files ({len(core_files)}):")
    for f in core_files:
        print(f"      - {f}")

    print(f"   📅 Cohort files ({len(cohort_files)}):")
    for f in cohort_files:
        print(f"      - {f}")

    # Run with all files
    generator.run(available_files)

    return generator

# Run the system
if __name__ == "__main__":
    generator = main()

    print(f"\n🚀 SYSTEM READY FOR 12 FILES!")
    print("📁 To add more files, place them in the directory:")
    print("   📚 Core: DEPT.xlsx (e.g., BSIT.xlsx, BSAI.xlsx)")
    print("   📅 Cohort: cohort_DEPT_X.X.xlsx")
    print("🎯 System will auto-detect and process all files!")