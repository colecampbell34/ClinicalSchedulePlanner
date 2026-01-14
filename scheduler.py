import pandas as pd
import random
from dataclasses import dataclass, field
from typing import List, Dict
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Border, Side

# ==================== DATA CLASSES ====================

@dataclass
class Site:
    site_id: float
    name: str
    is_general: bool
    is_cardiac: bool
    is_dual: bool
    high_acuity: bool
    out_of_town: bool
    only_may_to_dec: bool
    capacity: int

@dataclass
class Clinic:
    clinic_id: float
    name: str
    days_per_week: int
    out_of_town: bool
    only_may_to_dec: bool

@dataclass
class Student:
    student_id: str
    last_name: str
    first_name: str
    program: str
    choices: List[str]
    accommodation_sites: List[str]

    cardiac_weeks_completed: int = 0
    clinic_blocks: int = 0
    has_had_ob_exposure: bool = False

    # Track if this student is authorized for a Dual Chain
    is_chain_starter: bool = False 

    placements: Dict[str, str] = field(default_factory=dict)
    assigned_sites: List[str] = field(default_factory=list)

# ==================== SCHEDULER ====================

class Scheduler:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.students = []
        self.sites = []
        self.clinics = []

        self.cardiac_names = set()
        self.clinic_names = set()
        self.general_names = set()
        self.dual_site_names = set() 

        # (Student_ID, Block) -> "Red" | "Blue" | "Green" | "Orange" | "Purple" | "Warn"
        self.placement_intent = {}

        # Tracking
        self.site_slots = {b: defaultdict(list) for b in BLOCKS}
        self.clinic_days_used = {b: defaultdict(int) for b in BLOCKS}
        self.site_name_usage = {b: defaultdict(int) for b in BLOCKS}
        
        # QUOTA TRACKER
        self.dual_chains_assigned = 0
        self.MAX_DUAL_CHAINS = 5

    # ==================== LOAD ====================

    def load_data(self):
        try:
            students_df = pd.read_excel(self.excel_path, sheet_name="Students")
            sites_df = pd.read_excel(self.excel_path, sheet_name="Sites")
            clinics_df = pd.read_excel(self.excel_path, sheet_name="Clinics")
        except FileNotFoundError:
            print(f"Error: Could not find file '{self.excel_path}'")
            return

        for _, r in students_df.iterrows():
            choices = [
                str(r[f"choice_{i}"]).strip()
                for i in range(1, 6)
                if f"choice_{i}" in r and pd.notna(r[f"choice_{i}"])
            ]

            acc = []
            if pd.notna(r.get("accommodation_sites")):
                acc = [s.strip() for s in str(r["accommodation_sites"]).split(",")]

            self.students.append(Student(
                student_id=str(r["student_id"]),
                last_name=str(r["last_name"]),
                first_name=str(r["first_name"]),
                program=str(r["program"]),
                choices=choices,
                accommodation_sites=acc
            ))

        for _, r in sites_df.iterrows():
            s_name = str(r["site_name"]).strip()
            is_cardiac = bool(r["is_cardiac"])
            is_general = bool(r["is_general"])
            is_dual = bool(r["is_dual"])

            if is_cardiac: self.cardiac_names.add(s_name)
            if is_general: self.general_names.add(s_name)
            if is_dual: self.dual_site_names.add(s_name)

            self.sites.append(Site(
                site_id=r["site_id"],
                name=s_name,
                is_general=is_general,
                is_cardiac=is_cardiac,
                is_dual=is_dual,
                high_acuity=bool(r["high_acuity"]),
                out_of_town=bool(r["out_of_town"]),
                only_may_to_dec=bool(r["only_may_to_dec"]),
                capacity=int(r["capacity"])
            ))

        for _, r in clinics_df.iterrows():
            c_name = str(r["clinic_name"]).strip()
            self.clinic_names.add(c_name)
            self.clinics.append(Clinic(
                clinic_id=r["clinic_id"],
                name=c_name,
                days_per_week=int(r["days_per_week"]),
                out_of_town=bool(r["out_of_town"]),
                only_may_to_dec=bool(r["only_may_to_dec"])
            ))

    # ==================== SCHEDULING ====================

    def schedule_all(self):
        self.reset()
        for block in BLOCKS:
            print(f"Scheduling Block {block}...")
            self.schedule_block(block)
        
        self.apply_color_overrides()
        self.validate()
        self.print_choice_stats()
        return self.generate_output()

    def get_match_count(self, s):
        assigned_names_this_block = set()
        for b_key, p_val in s.placements.items():
            clean = p_val.split("(")[0].strip()
            assigned_names_this_block.add(clean)
        
        count = 0
        for choice in s.choices:
            if choice in assigned_names_this_block:
                count += 1
        return count

    def schedule_block(self, block):
        active = [s for s in self.students if self.is_active(s, block)]
        
        # --- 1. GUARANTEED CHOICE PHASE ---
        zero_match_students = [
            s for s in active 
            if self.get_match_count(s) == 0 and block not in s.placements and len(s.choices) > 0
        ]
        
        def zero_match_sort_key(s):
            bucket = 0 
            if block in ["4A", "4B"]:
                if s.program == "Cardiac": bucket += 1
                elif s.program == "Dual" and s.cardiac_weeks_completed < 12: bucket += 2
                else: bucket += 4
            else:
                if s.program == "Dual": 
                    # If behind schedule, boost priority
                    if block == "2B" and s.cardiac_weeks_completed == 0: bucket += 1
                    elif 0 < s.cardiac_weeks_completed < 12: bucket += 2
                    else: bucket += 10 
                else: bucket += 50
            return bucket, random.random()

        random.shuffle(zero_match_students)
        zero_match_students.sort(key=zero_match_sort_key)

        for s in zero_match_students:
            if block in s.placements: continue 
            
            dual_clinic_ban = (s.program == "Dual" and block in ["2A","2B","3A","3B"])
            
            placed = self.place_student(s, block, force_no_repeat=True, 
                                        is_zero_match_priority=True, only_choices=True, 
                                        ban_clinics=dual_clinic_ban)
            if not placed:
                placed = self.place_student(s, block, force_no_repeat=False, 
                                            is_zero_match_priority=True, only_choices=True,
                                            ban_clinics=dual_clinic_ban)
        
        # --- 2. REGULAR SCHEDULING ---
        remaining_active = [s for s in active if block not in s.placements]

        idx = BLOCKS.index(block)
        prev_block = BLOCKS[idx-1] if idx > 0 else None
        
        def regular_sort_key(s):
            must_continue = False
            if s.program == "Dual" and block in ["2B", "3B", "4B"] and s.is_chain_starter:
                must_continue = True
            
            bucket = 0 if must_continue else 100
            
            if block in ["4A", "4B"]:
                if s.program == "Cardiac": bucket += 1
                elif s.program == "Dual" and s.cardiac_weeks_completed < 12: bucket += 2
                else: bucket += 4
            else:
                if s.program == "Dual": 
                    # 2B Priority: 0 weeks -> High Priority
                    if block == "2B" and s.cardiac_weeks_completed == 0:
                        bucket += 5
                    # Finisher Priority:
                    elif 0 < s.cardiac_weeks_completed < 12:
                        bucket += 5
                    else:
                        bucket += 10
                else: bucket += 50
            
            matches = self.get_match_count(s)
            return bucket, matches, random.random()

        random.shuffle(remaining_active)
        remaining_active.sort(key=regular_sort_key)

        # 2a. Clinic attempts for General (High Priority)
        for s in remaining_active:
            if s.program == "General" and block not in s.placements:
                if block == "3B" and s.clinic_blocks >= 3: continue 
                if s.clinic_blocks < 3: 
                    prob = 0.5
                    if block in ["2A", "2B"]: prob = 0.90 
                    elif block in ["3A", "3B"] and s.clinic_blocks == 0: prob = 0.95
                    if random.random() < prob: 
                        self.try_clinic(s, block, force_no_repeat=True)

        # 2b. Strict Phase
        for s in remaining_active:
            if block in s.placements: continue
            placed = self.place_student(s, block, force_no_repeat=True)
            if not placed and s.program == "General" and s.clinic_blocks < 3:
                 placed = self.try_clinic(s, block, force_no_repeat=True)

        # 2c. Fallback Phase
        for s in remaining_active:
            if block in s.placements: continue
            if s.program == "General":
                if block == "3B" and s.clinic_blocks >= 3: continue
                if s.clinic_blocks < 3 and random.random() < 0.5:
                    self.try_clinic(s, block, force_no_repeat=False)

        for s in remaining_active:
            if block in s.placements: continue
            placed = self.place_student(s, block, force_no_repeat=False)
            if not placed and s.program == "General" and s.clinic_blocks < 3:
                 placed = self.try_clinic(s, block, force_no_repeat=False)
            
            if not placed:
                self.emergency_place(s, block)

    # ==================== COLOR OVERRIDES ====================

    def apply_color_overrides(self):
        for s in self.students:
            for a_block, b_block in [("2A", "2B"), ("3A", "3B"), ("4A", "4B")]:
                if a_block in s.placements and b_block in s.placements:
                    p1 = s.placements[a_block].split("(")[0].strip()
                    p2 = s.placements[b_block].split("(")[0].strip()
                    if p1 == p2 and p1 in self.dual_site_names:
                        self.placement_intent[(s.student_id, a_block)] = "Orange"
                        self.placement_intent[(s.student_id, b_block)] = "Orange"

            for block, placement in s.placements.items():
                p_clean = placement.split("(")[0].strip()
                if "Women" in p_clean or "BCWH" in p_clean:
                    self.placement_intent[(s.student_id, block)] = "Purple"

    # ==================== PLACEMENT LOGIC ====================

    def place_student(self, s, block, force_no_repeat=False, is_zero_match_priority=False, only_choices=False, ban_clinics=False):
        if s.program == "Cardiac":
            return self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                 is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

        if s.program == "Dual":
            # --- 1. CONTINUITY ---
            if block in ["2B", "3B", "4B"] and s.is_chain_starter:
                prev_block = BLOCKS[BLOCKS.index(block)-1]
                prev_site = s.placements.get(prev_block, "")
                if "(" in prev_site: prev_site = prev_site.split("(")[0].strip()
                
                target_site = next((x for x in self.sites if x.name == prev_site), None)
                if target_site:
                    if self.site_name_usage[block][target_site.name] < target_site.capacity:
                        slot_idx = self.find_slot(target_site, block, 5)
                        if slot_idx is not None:
                            self.site_slots[block][target_site.site_id][slot_idx] -= 5
                            self.site_name_usage[block][target_site.name] += 1
                            self.assign_site(s, block, target_site, is_cardiac_intent=False) 
                            s.is_chain_starter = False
                            return True
            
            # --- 2. 4A/4B LOGIC ---
            if block in ["4A", "4B"]:
                if ban_clinics: # Safety check
                    if s.cardiac_weeks_completed >= 12:
                        return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                             is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)
                    else:
                        return self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                             is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

                # DONE CARDIAC -> Clinic Priority
                if s.cardiac_weeks_completed >= 12:
                    if self.try_clinic(s, block, force_no_repeat=force_no_repeat, 
                                      is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    if self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                     is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    return False 
                else:
                    # NEED CARDIAC
                    if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                     is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    return self.try_clinic(s, block, force_no_repeat=force_no_repeat, 
                                          is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

            # --- 3. 2A-3B LOGIC ---
            if s.cardiac_weeks_completed < 13:
                
                # Check conflicts
                if self.is_same_number_cardiac_conflict(s, block):
                    return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                         is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

                # STRANDED / IN-PROGRESS: Force Cardiac
                if s.cardiac_weeks_completed > 0:
                    if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                     allow_new_dual_start=True, is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    return False

                # 0 Weeks: LOAD BALANCER
                
                # BLOCK 2A: 50/50 Split
                if block == "2A":
                    prefer_general = (random.random() < 0.5)
                    if prefer_general:
                        if self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                         is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                        if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                         allow_new_dual_start=True, is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    else:
                        if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                         allow_new_dual_start=True, is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                        if self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                         is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    return False

                # BLOCK 2B: Catch-up if missed 2A
                if block == "2B" and s.cardiac_weeks_completed == 0:
                    # FORCE CARDIAC (They chose Tails in 2A, now they must pay with Heads)
                    if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                     allow_new_dual_start=True, is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    # If no cardiac, they are in trouble for 3A/3B but allow general
                    return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                         is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

                # BLOCK 3A/3B: If 0 weeks, PANIC PRIORITY
                if block in ["3A", "3B"] and s.cardiac_weeks_completed == 0:
                    if self.try_site(s, block, cardiac=True, force_no_repeat=force_no_repeat, 
                                     allow_new_dual_start=True, is_zero_match_priority=is_zero_match_priority, only_choices=only_choices): return True
                    return False

                # Default fallback
                return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                     is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

            # Done Cardiac, prefer General
            return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                 is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)

        if s.program == "General":
            return self.try_site(s, block, general=True, force_no_repeat=force_no_repeat, 
                                 is_zero_match_priority=is_zero_match_priority, only_choices=only_choices)
        return False

    def try_site(self, s, block, cardiac=False, general=False, force_no_repeat=False, allow_new_dual_start=True,
                 is_zero_match_priority=False, only_choices=False):
        
        candidates = []
        if cardiac: candidates.extend([x for x in self.sites if x.is_cardiac])
        if general: candidates.extend([x for x in self.sites if x.is_general])
        
        unique_map = {x.site_id: x for x in candidates}
        candidates = list(unique_map.values())
        
        # REMOVE SITES UNAVAILABLE IN SPRING
        if block in ["2A", "2B", "3A"]:
            candidates = [x for x in candidates if not x.only_may_to_dec]

        last_site_name = ""
        idx = BLOCKS.index(block)
        if idx > 0:
            prev_block = BLOCKS[idx-1]
            last_site_name = s.placements.get(prev_block, "")
            if "(" in last_site_name: last_site_name = last_site_name.split("(")[0].strip()

        candidates = [x for x in candidates if self.allow_site(s, x)]

        filtered = []
        for x in candidates:
            is_continuity = (x.name == last_site_name)
            if x.is_dual and s.program == "Dual" and not is_continuity and not allow_new_dual_start:
                continue 

            if only_choices and x.name not in s.choices: 
                continue

            # Math Guard for Cardiac
            if s.program == "Dual" and cardiac:
                is_likely_cardiac_intent = False
                if x.is_cardiac:
                    if not x.is_general: is_likely_cardiac_intent = True
                    elif cardiac: is_likely_cardiac_intent = True
                
                if is_likely_cardiac_intent:
                    potential_total = s.cardiac_weeks_completed + BLOCK_WEEKS[block]
                    if potential_total > 13:
                        continue 

            if force_no_repeat:
                if x.name in s.assigned_sites: continue
                filtered.append(x)
            else:
                filtered.append(x)
        
        candidates = filtered

        current_matches = self.get_match_count(s)
        ignore_choices_for_score = (current_matches >= 2) and not is_zero_match_priority

        def site_score(site_for_scoring):
            score = 0
            
            if not is_zero_match_priority:
                score += (current_matches * 10000)

            is_bcwh = "Women" in site_for_scoring.name or "BCWH" in site_for_scoring.name
            if is_bcwh:
                if s.program == "Dual": score -= 5000
                else: score += 5000
            
            if site_for_scoring.name in s.assigned_sites: score += 100000

            if s.program == "Dual" and site_for_scoring.is_dual:
                if site_for_scoring.name != last_site_name:
                    is_top_choice = False
                    try:
                        s.choices.index(site_for_scoring.name)
                        is_top_choice = True
                    except ValueError: pass

                    if is_top_choice and not ignore_choices_for_score and self.dual_chains_assigned < self.MAX_DUAL_CHAINS:
                        score -= 500 
                    else:
                        score += 1000 

            if ignore_choices_for_score:
                score += 500 
            else:
                try: 
                    score += s.choices.index(site_for_scoring.name)
                except ValueError: 
                    score += 500 
            
            if site_for_scoring.out_of_town and site_for_scoring.name not in s.accommodation_sites: score += 500
            if not cardiac and site_for_scoring.is_cardiac: score += 1000

            return score, random.random()

        candidates.sort(key=site_score)

        for site in candidates:
            current_usage = self.site_name_usage[block][site.name]
            if current_usage >= site.capacity: continue

            slot_idx = self.find_slot(site, block, 5)
            if slot_idx is not None:
                self.site_slots[block][site.site_id][slot_idx] -= 5
                self.site_name_usage[block][site.name] += 1
                
                if s.program == "Dual" and site.is_dual and block in ["2A", "3A", "4A"]:
                    if self.dual_chains_assigned < self.MAX_DUAL_CHAINS:
                        s.is_chain_starter = True
                        self.dual_chains_assigned += 1
                
                is_actually_cardiac = False
                if site.is_cardiac and not site.is_general:
                    is_actually_cardiac = True
                elif site.is_cardiac and site.is_general:
                    if cardiac and not general: is_actually_cardiac = True
                    elif general and not cardiac: is_actually_cardiac = False
                    else:
                        if s.program == "Cardiac": is_actually_cardiac = True
                        elif s.program == "Dual" and s.cardiac_weeks_completed < 13: is_actually_cardiac = True
                        else: is_actually_cardiac = False

                self.assign_site(s, block, site, is_actually_cardiac)
                return True
        return False

    def try_clinic(self, s, block, force_no_repeat=False, is_zero_match_priority=False, only_choices=False, ban_clinics=False):
        if ban_clinics: return False

        current_matches = self.get_match_count(s)
        ignore_choices_for_score = (current_matches >= 2) and not is_zero_match_priority

        def clinic_score(c):
            score = 0
            if not is_zero_match_priority:
                score += (current_matches * 10000)

            if c.name in s.assigned_sites: score += 100000
            
            if ignore_choices_for_score:
                score += 500
            else:
                try: 
                    score += s.choices.index(c.name)
                except ValueError: 
                    score += 500
            
            if c.out_of_town and c.name not in s.accommodation_sites: score += 500
            return score, random.random()

        candidates = self.clinics
        # REMOVE CLINICS UNAVAILABLE IN SPRING
        if block in ["2A", "2B", "3A"]:
            candidates = [c for c in candidates if not c.only_may_to_dec]

        if only_choices:
            candidates = [c for c in candidates if c.name in s.choices]

        if force_no_repeat:
             candidates = [c for c in candidates if c.name not in s.assigned_sites]
        sorted_clinics = sorted(candidates, key=clinic_score)

        for c in sorted_clinics:
            if self.clinic_days_used[block][c.clinic_id] > 0: continue
            self.assign_clinic(s, block, c)
            return True
        return False

    def emergency_place(self, s, block):
        if s.program == "General":
             if self.try_clinic(s, block, force_no_repeat=False):
                 self.placement_intent[(s.student_id, block)] = "Green"
                 return
             if self.try_site(s, block, cardiac=True, general=True, force_no_repeat=False):
                 self.placement_intent[(s.student_id, block)] = "Blue"
                 return

        if s.program == "Dual":
            if s.cardiac_weeks_completed >= 12:
                if self.try_site(s, block, general=True, force_no_repeat=False):
                    self.placement_intent[(s.student_id, block)] = "Blue"
                    return
                if self.try_clinic(s, block, force_no_repeat=False):
                    self.placement_intent[(s.student_id, block)] = "Green"
                    return
            else:
                if self.try_site(s, block, cardiac=True, force_no_repeat=False):
                    self.placement_intent[(s.student_id, block)] = "Red"
                    return
                if self.try_site(s, block, general=True, force_no_repeat=False):
                    self.placement_intent[(s.student_id, block)] = "Blue"
                    return
                if self.try_clinic(s, block, force_no_repeat=False):
                    self.placement_intent[(s.student_id, block)] = "Green"
                    return

        s.placements[block] = "UNASSIGNED"
        self.placement_intent[(s.student_id, block)] = "Warn"

    # ==================== HELPERS ====================

    def is_same_number_cardiac_conflict(self, s, block):
        idx = BLOCKS.index(block)
        if idx == 0: return False
        prev_block = BLOCKS[idx - 1]
        
        prev_intent = self.placement_intent.get((s.student_id, prev_block))
        
        if prev_intent == "Red":
            if block[0] == prev_block[0]: 
                return True
        return False

    def find_slot(self, site, block, days_needed):
        slots = self.site_slots[block][site.site_id]
        for i, remaining in enumerate(slots):
            if remaining >= days_needed: return i
        return None

    def assign_site(self, s, block, site, is_cardiac_intent):
        s.placements[block] = site.name
        s.assigned_sites.append(site.name)
        
        if is_cardiac_intent:
            self.placement_intent[(s.student_id, block)] = "Red"
            s.cardiac_weeks_completed += BLOCK_WEEKS[block]
        else:
            self.placement_intent[(s.student_id, block)] = "Blue"
        
        if not is_cardiac_intent and site.name != "Vancouver General":
            s.has_had_ob_exposure = True

    def assign_clinic(self, s, block, clinic):
        self.clinic_days_used[block][clinic.clinic_id] += clinic.days_per_week
        s.placements[block] = f"{clinic.name} ({clinic.days_per_week}d)"
        s.assigned_sites.append(clinic.name)
        s.clinic_blocks += 1
        s.has_had_ob_exposure = True
        self.placement_intent[(s.student_id, block)] = "Green"

    def allow_site(self, s, site):
        if site.name == "Vancouver General" and not s.has_had_ob_exposure: return False
        return True

    def is_active(self, s, block):
        if s.program == "General" and block not in ["2A","2B","3A","3B"]: return False
        if s.program == "Dual" and block == "Summer": return False
        if s.program == "Cardiac" and block in ["2A","2B","3A","3B"]: return False
        return True

    def reset(self):
        self.placement_intent.clear()
        self.dual_chains_assigned = 0
        for s in self.students:
            s.cardiac_weeks_completed = 0
            s.clinic_blocks = 0
            s.is_chain_starter = False
            s.has_had_ob_exposure = False
            s.placements.clear()
            s.assigned_sites.clear()
        for b in BLOCKS:
            self.clinic_days_used[b].clear()
            self.site_name_usage[b].clear() 
            self.site_slots[b].clear()
            for site in self.sites:
                self.site_slots[b][site.site_id] = [5] * site.capacity

    def print_choice_stats(self):
        print("\n" + "="*40)
        print("       STUDENT CHOICE SATISFACTION       ")
        print("="*40)
        
        total_students = len(self.students)
        total_matches = 0
        students_with_0 = 0
        
        print(f"{'Name':<25} | {'Matches':<8} | {'Matched Choices'}")
        print("-" * 65)
        
        for s in self.students:
            matches = self.get_match_count(s)
            matched_list = []
            for p in s.placements.values():
                clean = p.split("(")[0].strip()
                if clean in s.choices: matched_list.append(clean)
            
            total_matches += matches
            if matches == 0:
                students_with_0 += 1
                
            print(f"{s.first_name} {s.last_name:<15} | {matches}/5     | {', '.join(set(matched_list))}")

        avg = total_matches / total_students if total_students > 0 else 0
        print("-" * 65)
        print(f"Total Students: {total_students}")
        print(f"Average Choices Met: {avg:.2f} / 5.0")
        print(f"Students with 0 matches: {students_with_0}")
        print("="*40 + "\n")

    def validate(self):
        print("\n--- VALIDATION ---")
        unassigned_count = 0
        for s in self.students:
            if s.program == "Dual":
                if s.cardiac_weeks_completed < 12:
                    print(f"⚠️ {s.first_name} (Dual) only has {s.cardiac_weeks_completed} cardiac weeks.")
            
            site_counts = pd.Series(s.assigned_sites).value_counts()
            repeats = site_counts[site_counts > 1]
            if not repeats.empty:
                pass 

            for b in BLOCKS:
                if self.is_active(s, b):
                    if b not in s.placements or "UNASSIGNED" in s.placements[b]:
                        print(f"⚠️ {s.first_name} unassigned in {b}")
                        unassigned_count += 1
        
        if unassigned_count == 0:
            print("✅ All active slots filled.")
        else:
            print(f"❌ {unassigned_count} unassigned slots found.")

    def generate_output(self):
        order = {"General": 0, "Dual": 1, "Cardiac": 2}
        rows = []
        for s in sorted(self.students, key=lambda x: order.get(x.program, 3)):
            row = {
                "student_id": s.student_id,
                "last_name": s.last_name,
                "first_name": s.first_name,
                "program": s.program
            }
            for b in BLOCKS:
                row[b] = s.placements.get(b, "")
            rows.append(row)
        return pd.DataFrame(rows)

# ==================== EXCEL COLORING ====================

def apply_formatting(filename, scheduler):
    print(f"Applying color coding to {filename}...")
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    except Exception as e:
        print(f"Error opening workbook for styling: {e}")
        return

    # Styles
    fill_cardiac = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Red
    fill_clinic = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Green
    fill_general = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid") # Blue
    fill_dual = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")    # Peachy Orange
    fill_bcwh = PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")    # Light Lavender
    fill_warn = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")    # Yellow
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # Header Mapping
    header_row = ws[1]
    block_col_indices = {}
    for cell in header_row:
        cell.border = thin_border
        if cell.value in BLOCKS:
            block_col_indices[cell.col_idx] = cell.value

    # Row Iteration
    for row in ws.iter_rows(min_row=2):
        student_id = str(row[0].value) 
        
        for cell in row:
            cell.border = thin_border
            
            if cell.col_idx in block_col_indices:
                block = block_col_indices[cell.col_idx]
                
                # Check Intent
                intent = scheduler.placement_intent.get((student_id, block))
                
                if intent == "Red":
                    cell.fill = fill_cardiac
                elif intent == "Blue":
                    cell.fill = fill_general
                elif intent == "Green":
                    cell.fill = fill_clinic
                elif intent == "Orange":
                    cell.fill = fill_dual
                elif intent == "Purple":
                    cell.fill = fill_bcwh
                elif intent == "Warn":
                    cell.fill = fill_warn
                
                if "UNASSIGNED" in str(cell.value):
                    cell.fill = fill_warn

    wb.save(filename)
    print("✅ Color coding and borders applied.")

# ==================== CONSTANTS ====================

BLOCKS = ["2A", "2B", "3A", "3B", "Summer", "4A", "4B"]
BLOCK_WEEKS = {"2A":7,"2B":6,"3A":6,"3B":7,"Summer":7,"4A":7,"4B":6}

# ==================== MAIN ====================

def main():
    scheduler = Scheduler("Student_Site_Data.xlsx")
    scheduler.load_data()
    if not scheduler.students:
        print("No students loaded. Check Excel file.")
        return
    
    df = scheduler.schedule_all()
    output_file = "Generated_Schedule.xlsx"
    df.to_excel(output_file, index=False)
    
    apply_formatting(output_file, scheduler)
    print(f"Schedule saved to {output_file}")

if __name__ == "__main__":
    main()