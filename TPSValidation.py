"""
Script that automatically calculates and creates a text document to paste into the TPS validation report. 
Puts the dose values to a file called "doses." This script may require setting up the file_name variable to point to a user accessible file location.
If the file already exists it will create a new file with an incrementing number added.
Check the section below headed "DEFAULT LOCAL ENERGIES" to adjust

not sure if functional
"""

file_name = '\\\\Client\\Z$\\RayStation\\Scripts\\doses.txt'

from connect import *
import itertools
from datetime import datetime
import os

case = get_current("Case")
examination = get_current("Examination")
open_field_plan = case.TreatmentPlans["TG53 6x"]
edw_plan = case.TreatmentPlans["TG53 6x edw"]

# SETTINGS
MACHINE_HAS_EDW = True
MACHINE_HAS_HARD_WEDGES = False
MUS = 1000
DOSE_ALGO = "CCDose"

# DEFAULT LOCAL ENERGIES
local_energies = ["6X", "10X", "18X", "6 FFF", "10 FFF"]
all_field_sizes = ["5x5", "10x10", "10x10-5x5mlc", "20x20", "30x30", "8x8"]

# DETAILS FOR GENERATING CORRECT ORDER OF DOSES FOR EXCEL
report_energies = ["6X", "8X", "10X", "15X", "18X", "6 FFF", "10 FFF"]
print_out_structure = ["Open Field", "Off-Axis open", "EDW", "Off-Axis EDW", "Hard Wedge"]
open_field_sizes = ["5x5", "10x10", "10x10-5x5mlc", "20x20", "30x30"]
off_axis_field_sizes = ["10x10", "20x20", "30x30", "8x8 Offset"]
edw_fields = ["30 in", "30 out", "60 in", "60 out"]
hard_wedge_fields = ["W15", "W30", "W45", "W60"]
depths = ["10 cm", "20 cm"]
non_fff_fields = print_out_structure[2:]


# CLASS STRUCTURES FOR PRINTING OUT TO FILE TO PASTE INTO REPORT EXCEL

class RowItem:
    def __init__(self, depth, field, energy):
        self.depth = depth
        self.field_size = field
        self.energy = energy
        self.dose = ""

    def get_doses(self):
        return f"{self.dose}\n"


class Section:
    def __init__(self, name):
        self.name = name
        self.items = []

    def get_doses(self):
        col = ""
        for item in self.items:
            if self.name == "FullReport":
                col += f"TPS Dose (cGy)\n"
            col += item.get_doses()
        if self.name in print_out_structure:
            col += "\n"

        return col


class SubSection(Section):
    def __init__(self, name):
        Section.__init__(self, name)
        for energy in report_energies:
            if (self.name in non_fff_fields and "FFF" not in energy) or not self.name in non_fff_fields:
                new_section = SubSubSection(self.name, energy)
                self.items.append(new_section)

    def add_dose(self, dose):
        for section in self.items:
            if section.name == dose[1]:
                section.add_dose(dose)
                break


class SubSubSection(Section):
    def __init__(self, parent, name):
        Section.__init__(self, name)
        self.parent_name = parent
        if self.parent_name == "Open Field":
            for depth in depths:
                for field in open_field_sizes:
                    self.items.append(RowItem(depth, field, self.name))
        elif self.parent_name == "Off-Axis open":
            for field in off_axis_field_sizes:
                self.items.append(RowItem("10 cm", field, self.name))
        elif self.parent_name in ["EDW", "Off-Axis EDW"]:
            for field in edw_fields:
                self.items.append(RowItem("10 cm", field, self.name))
        elif self.parent_name == "Hard Wedge":
            for field in hard_wedge_fields:
                self.items.append(RowItem("10 cm", field, self.name))

    def add_dose(self, dose):
        for row in self.items:
            if row.depth == dose[2] and row.field_size == dose[3]:
                row.dose = str(dose[4])
                break


class FullReport(Section):
    def __init__(self, name="FullReport"):
        Section.__init__(self, name)
        for section_name in print_out_structure:
            self.items.append(SubSection(section_name))

    def add_doses(self, doses):
        for dose in doses:
            for section in self.items:
                if section.name == dose[0]:
                    section.add_dose(dose)
                    break


# RAYSTATION LOGIC FOR CALCULATING AND EXTRACTING DOSES


def get_open_field_doses(energy):
    doses = []
    for depth in depths:
        for open_field in open_field_sizes:
            beam = open_field_plan.BeamSets[0].Beams[open_field]
            dose = get_dose_at(open_field_plan, beam, depth)
            new_dose_entry = ("Open Field", energy, depth, open_field, dose)
            doses.append(new_dose_entry)

    poi_dict = {"10x10": "Off Axis 1", "20x20": "Off Axis 2", "30x30": "Off Axis 3", "8x8 Offset": "Off Axis 4"}
    for off_axis in off_axis_field_sizes:
        beam = open_field_plan.BeamSets[0].Beams[off_axis]
        dose = get_dose_at(open_field_plan, beam, poi_dict[off_axis])
        new_dose_entry = ("Off-Axis open", energy, "10 cm", off_axis, dose)
        doses.append(new_dose_entry)
    print(doses)
    return doses


def get_edw_doses(energy):
    doses = []
    for FS in ["20x20", "8x8"]:
        for EDW in edw_fields:
            beam_name = FS + " " + EDW
            beam = edw_plan.BeamSets[0].Beams[beam_name]
            point = "10 cm" if FS == "20x20" else "Off Axis 4"
            section_name = "EDW" if FS == "20x20" else "Off-Axis EDW"
            dose = get_dose_at(edw_plan, beam, point)
            new_dose_entry = (section_name, energy, "10 cm", EDW, dose)
            doses.append(new_dose_entry)
    print(doses)
    return doses


def set_energies_in_plan(energy, plan):
    energy_id = energy.strip("X")
    beams = plan.BeamSets[0].Beams
    for beam in beams:
        beam.BeamQualityId = energy_id
        beam.BeamMU = MUS


def get_dose_at(plan, beam, point):
    # point will be the name of the point in the POI list, ie "isocenter", "20 cm", "1.4 cm" etc.
    index = beam.Number - 1
    BeamDose = plan.BeamSets[0].FractionDose.BeamDoses[index]
    poi = case.PatientModel.StructureSets[examination.Name].PoiGeometries[point]
    coords = {"x": poi.Point.x, "y": poi.Point.y, "z": poi.Point.z}
    dose = BeamDose.InterpolateDoseInPoint(Point=coords,
                                           PointFrameOfReference=examination.EquipmentInfo.FrameOfReference)
    return dose * 100 / MUS


def main():
    global file_name
    new_report = FullReport()
    for energy in local_energies:
        set_energies_in_plan(energy, open_field_plan)
        open_field_plan.BeamSets[0].ComputeDose(DoseAlgorithm=DOSE_ALGO, ForceRecompute=1)
        new_report.add_doses(get_open_field_doses(energy))
        if "FFF" not in energy and MACHINE_HAS_EDW:
            set_energies_in_plan(energy, edw_plan)
            edw_plan.BeamSets[0].ComputeDose(DoseAlgorithm=DOSE_ALGO, ForceRecompute=1)
            new_report.add_doses(get_edw_doses(energy))
        if "FFF" not in energy and MACHINE_HAS_HARD_WEDGES:
            return Error('hard wedges not implemented in script')
    output = new_report.get_doses()
    print("File sample...")
    print(output)
    orig_file_name = file_name
    i = 1
    while os.path.isfile(file_name):
        file_name = file_name.strip(".txt") + str(i) + ".txt"
        i += 1
    with open(file_name, "a") as dose_file:
        dose_file.write(datetime.now().strftime("%D %H:%M:%S") + "\n")
        dose_file.write(output)


main()
# test_report = FullReport()
# for subsection in test_report.items:
#    print(subsection.name)
#    for energy in subsection.items:
#        print("      " + energy.name)
#        for field in energy.items:
#            print("                   " + field.field_size)
