"""
This attached script is provided as a tool and not as a RaySearch endorsed script for
clinical use.  The use of this script in its whole or in part is done so
without any guarantees of accuracy or expected outcomes. Please verify all results. Further,
per the Raystation Instructions for Use, ALL scripts MUST be verified by the user prior to
clinical use.

Tomo_Check.py
RS 9B, 10, 11A
v1.3 20211014

This script will run on all beams on all beam sets in the plan
and check for Tomotherapy limits applied to
+the localization point shift to isocenter
+potential bore collisions with calculation structures
+isocenter placement relative to image stack center
then a pop-up window will report the results listed by beam set
all margins/tolerances are found in VARIABLE DEFINITION
"""

from connect import *

import clr

patient = get_current('Patient')
case = get_current('Case')
plan = get_current('Plan')

# beam_set = get_current("BeamSet")
examination = get_current("Examination")
structure_set = case.PatientModel.StructureSets[examination]
name_of_couch = "Upper pallet"
couch_structure = structure_set.RoiGeometries[name_of_couch]
couch_z = couch_structure.GetCenterOfRoi().z

################################
# VARIABLE DEFINITION
# check for iso w/in 10cm of imageStack center variables
iso_shift = 10.0 # cm
# check for intersection with BORE variables
bore_diameter = 85.0 # cm
# bore_length = 135.0 # cm
bore_length = 2 * examination.Series[0].ImageStack.SlicePositions[-1] # cm
collision_margin = 2.0 # cm
# check for lateral shift to iso variables
lateral_shift_of_the_laser_offset = 2.5 # cm
# max height of iso from couch
max_vertical_distance_to_couch = 23.5 # cm

################################
# MAIN CODE
# output init
text = ''
iso_list = []
# check for iso w/in 10cm of imageStack center
ct_box = case.Examinations[0].Series[0].ImageStack.GetBoundingBox()
x_center = (ct_box[1].x + ct_box[0].x)/2 # DICOM +X, IEC +X
# y_center = (ct_box[1].y + ct_box[0].y)/2 # DICOM +Y, IEC -Z
# z_center = (ct_box[1].z + ct_box[0].z)/2 # DICOM +Z, IEC +Y

for beam_set in plan.BeamSets:
    text += ' \nBeamSet: {}\n'.format(beam_set.DicomPlanLabel)
    for beam in beam_set.Beams:
        iso_position = beam.Isocenter.Position
        print ("iso position is: {}".format(iso_position))
        print ("difference is: {}".format(iso_position.x - x_center))
        if abs(iso_position.x - x_center) >= iso_shift:
            print ('The distance to isocenter exceeeds the limit in beam {}'.format(beam.Name))
            text += 'Iso w/in 10cm of center of image set:\n\t---Limit EXCEEDED in beam {}---\n'.format(beam.Name)
        else:
            print ('Isocenter placement in ImageStack is OK in beam {}'.format(beam.Name))
            text += 'Iso w/in 10cm of center:\n\t+++OK for beam {}+++\n'.format(beam.Name)
            
        ###################################
        # check for vertical shift to iso
        # this is undesirable as the couch
        # cannot go any lower than ~25.5 cm
        iso_height_from_couch = abs(iso_position.z - couch_z)
        if iso_height_from_couch >= max_vertical_distance_to_couch:
            text += "Iso height is {} cm above the couch, this EXCEEDS the limit and the couch may not be able to go low enough".format(iso_height_from_couch)
        else:
            text += "Iso height of {} is OK and within vertical shift limits".format(iso_height_from_couch)

        ################################
        # check for intersection with BORE
        temp_bore_str = case.PatientModel.GetUniqueRoiName(DesiredName = "_temp_bor")
        temp_col_str = case.PatientModel.GetUniqueRoiName(DesiredName = "_temp_coll")
        temp_int_str = case.PatientModel.GetUniqueRoiName(DesiredName = "_temp_inter")
        support_rois = [temp_bore_str, temp_col_str, temp_int_str]
        # other_roi = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Name not in support_rois]
        calc_rois_types = ['Support', 'Fixation', 'Bolus', 'External']
        other_roi = [roi.Name for roi in case.PatientModel.RegionsOfInterest if roi.Type in calc_rois_types]

        with CompositeAction('Create ROIs'):
            temp_bore_roi = case.PatientModel.CreateRoi(Name=temp_bore_str, Color="Yellow",
                                                        Type="Organ", TissueName=None,
                                                        RbeCellTypeName=None, RoiMaterial=None)
            patient.SetRoiVisibility(RoiName = temp_bore_str, IsVisible = False)
            temp_bore_roi.CreateCylinderGeometry(Radius=bore_diameter/2.0, Axis={ 'x': 0, 'y': 0, 'z': 1 },
                                                 Length=bore_length,
                                                 Examination=examination,
                                                 Center={ 'x': iso_position.x, 'y': iso_position.y, 'z': iso_position.z },
                                                 Representation="TriangleMesh", VoxelSize=None)

            temp_collision_roi = case.PatientModel.CreateRoi(Name=temp_col_str, Color="255, 128, 0", Type="Organ",
                                                             TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
            temp_collision_roi.SetAlgebraExpression(
                ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [temp_bore_str],
                              'MarginSettings': { 'Type': "Contract", 'Superior': 0, 'Inferior': 0,
                                                  'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } },
                ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [temp_bore_str],
                              'MarginSettings': { 'Type': "Contract", 'Superior': 0, 'Inferior': 0,
                                                  'Anterior': collision_margin, 'Posterior': collision_margin,
                                                  'Right': collision_margin, 'Left': collision_margin } },
                ResultOperation="Subtraction",
                ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                       'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
            patient.SetRoiVisibility(RoiName = temp_col_str, IsVisible = False)
            temp_collision_roi.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")


            temp_intersection_roi = case.PatientModel.CreateRoi(Name=temp_int_str, Color="SaddleBrown", Type="Organ",
                                                              TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
            temp_intersection_roi.SetAlgebraExpression(
                ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [temp_col_str],
                              'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                                  'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } },
                ExpressionB={ 'Operation': "Union", 'SourceRoiNames': other_roi,
                              'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                                  'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } },
                ResultOperation="Intersection",
                ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                       'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })
            patient.SetRoiVisibility(RoiName = temp_int_str, IsVisible = False)
            temp_intersection_roi.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        if case.PatientModel.StructureSets[examination.Name].RoiGeometries[temp_int_str].HasContours():
            print ('+++++ COLLISION DETECTED +++++')
            text += 'Structures w/in {} cm of bore:\n\t---COLLISION DETECTED---\n'.format(collision_margin)
        else:
            print('No collision from beam {} isocenter'.format(beam.Name))
            text += 'Structures w/in {} cm of bore:\n\t+++No collision+++\n'.format(collision_margin)

        with CompositeAction('Cleanup ROIs'):
            for roi in support_rois:
                case.PatientModel.RegionsOfInterest[roi].DeleteRoi()

    ################################
    # check for lateral shift to iso
    try:
        loc_point = [poi.Name for poi in case.PatientModel.PointsOfInterest if poi.Type == 'LocalizationPoint'][0]
        loc_point_coords = case.PatientModel.StructureSets[examination.Name].PoiGeometries[loc_point].Point

        if abs(loc_point_coords.x - iso_position.x) >= lateral_shift_of_the_laser_offset:
            print('The shift to Isocenter from {} exceeds the limit.'.format(loc_point))
            text += 'lateral shift of the laser offset w/in {}cm of Iso:\n\t---EXCEEDS LIMIT---\n'.format(
                lateral_shift_of_the_laser_offset)
        else:
            print('lateral shift of the laser offset OK')
            text += 'lateral shift of the laser offset w/in {}cm of Iso:\n\t+++Shift OK+++\n'.format(
                lateral_shift_of_the_laser_offset)
    except:
        print ("No localization point defined")
        text += 'lateral shift of the laser offset w/in {}cm of Iso:\n\t---No localization point defined---\n'.format(
            lateral_shift_of_the_laser_offset)

print ('**** Checks Complete ****')
text += '\n\t**** Checks Complete ****\n'
################################

try:
    clr.AddReference('System.Windows.Forms')
    from System.Windows.Forms import MessageBox
    MessageBox.Show(text)
except:
    import ctypes  # An included library with Python install.
    ctypes.windll.user32.MessageBoxW(0, text, "Paddick & Homogeneity", 0)
