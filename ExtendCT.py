#ExtendCT.py
# use to copy CT slices in situations in which the calculation volume must extend beyond the scanned region. Do not extend the treatment area into extended slices.
from connect import *
import os
import pydicom

EXTEND_ON = "INFERIOR" # set to "INF" or "SUP"
EXTENSION_SIZE = "50" # in millimeters


patient = get_current("Patient")
patient_db = get_current('PatientDB')
exam = get_current("Examination")
case = get_current("Case")
series = exam.Series[0]
stack = series.ImageStack
SliceThickness = abs(series.ImageStack.SlicePositions[0] - series.ImageStack.SlicePositions[1])*10
number_of_slices_to_extend = int(EXTENSION_SIZE/SliceThickness)+1

tmp = '\\\\Client\\Z$\\RayStation\\Scripts\\tmp' #Change to something applicable in your environment.
os.chdir(tmp)

def ExtendCT():
    
    print("ExtendCT not implemented")

def LogWarning(error):
    try:
        jsonWarnings = json.loads(str(error))
        # If the json.loads() works then the script was stopped due to
        # a non-blocking warning.
        print ("WARNING! Export Aborted!")
        print ("Comment:")
        print (jsonWarnings["Comment"])
        print ("Warnings:")

        # Here the user can handle the warnings. Continue on known warnings,
        # stop on unknown warnings.
        for w in jsonWarnings["Warnings"]:
            print (w)
    except ValueError as error:
        print ("Error occurred. Could not export.")

# This prints the successful result log in an ordered way.
def LogCompleted(result):
    try:
        jsonWarnings = json.loads(str(result))
        print ("Completed!")
        print ("Comment:")
        print (jsonWarnings["Comment"])
        print ("Warnings:")
        for w in jsonWarnings["Warnings"]:
            print (w)
        print ("Export notifications:")
        # Export notifications is a list of notifications that the user should read.
        for w in jsonWarnings["Notifications"]:
            print (w)
    except ValueError as error:
        print ("Error reading completion messages.")


def ExportCurrentExamination():
    try:
        results = case.ScriptableDicomExport(ExportFolderPath = tmp,
                                             DicomFilter = "",
                                             IgnorePreConditionWarnings = False,
                                             Examinations= [exam.Name]
                                             )
        LogCompleted(result)

    except System.InvalidOperationException as error:
        LogWarning(error)
        results = case.ScriptableDicomExport(ExportFolderPath = tmp,
                                             DicomFilter = "",
                                             IgnorePreConditionWarnings = True,
                                             Examinations= [exam.Name]
                                             )
    except Exception as e:
        print("Except %s" % e)

def cleartmp():
    for f in os.listdir():
        os.remove(f)

def ImportModifiedExamination():
    patient_id = patient.PatientID
    matching_patients = patient_db.QueryPatientsFromPath(Path = tmp, SearchCriterias = {'PatientID' : patient_id})
    assert len(matching_patients) == 1, f"Found more than1 patient with ID {patient_id}"
    matching_patient = matching_patients[0]
    studies = patient_db.QueryStudiesFromPath(Path = tmp, SearchCriterias = matching_patient)
    series = []
    for study in studies:
        series += patient_db.QuerySeriesFromPath(Path = tmp, SearchCriterias = study)
    seriesToImport = [s for s in series if s['Modality'] == "CT"]
    warnings = patient.ImportDataFromPath(Path = tmp, SeriesOrInstances = seriesToImport, CaseName = case.CaseName)
    print(f"Warnings: {warnings}")



cleartmp()
ExportCurrentExamination()
ExtendCT()
ImportModifiedExamination()
patient.Save()