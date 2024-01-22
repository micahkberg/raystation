#For extraction of average doses to an ROI. Change name to name.py. Edit the path to the dose extraction file. Change machine name. Run from the RayStation scripting. 
from connect import *
from datetime import datetime

output = datetime.now().strftime("%D %H:%M:%S") +"\n"

patient = get_current("Patient")
case = patient.Cases[0]
#beam_set = get_current('BeamSet')
file_name = '\\\\Client\\Z$\\RayStation\\Scripts\\doses.txt' #Change to something applicable in your environment.

for plan in case.TreatmentPlans:
  output+=plan.Name+"\n" 
  if 'MLCi2' in plan.Name or True: #Can be removed, or edited to match the MLC
    #plan.BeamSets[0].SetMachineReference(MachineName = 'TrueBeam 12A' ) #Change to your machine name
    plan.BeamSets[0].ComputeDose(DoseAlgorithm='CCDose', ForceRecompute = 1) 
    #plan.BeamSets[0].AccurateDoseAlgorithm.MCStatisticalUncertaintyForFinalDose = 0.005 #0.001-0.1#
    #plan.BeamSets[0].ComputeDose(DoseAlgorithm='PhotonMonteCarlo', ForceRecompute = 1)
    beam_set = plan.BeamSets[0]
    fraction_dose = beam_set.FractionDose

    doses = []
    for idx, beam_dose in enumerate(fraction_dose.BeamDoses):
        planname = plan.Name
        roiDose = beam_dose.GetDoseStatistic(RoiName = 'LargeEvalVolume', DoseType = 'Average')
        beamname = beam_set.Beams[idx].Name
        doses.append([beamname, roiDose])
    
    for dose in doses:
      output+=f"{dose[0]}\t{dose[1]}\n"
    output+="\n"

    #with open(file_name, "a") as dose_file:
      # old print method
      #dose_file.write(plan.Name + " =[")
      #for d in doses:
      #  dose_file.write(str(d) + " ")
      #dose_file.write("];\n")

    patient.Save()

with open(file_name, "a") as dose_file:
  dose_file.write(output)
