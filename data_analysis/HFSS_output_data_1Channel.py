# ----------------------------------------------
# Script Recorded by ANSYS Electronics Desktop Version 2020.2.0
# 16:23:53  March 18, 2021
# ----------------------------------------------
import ScriptEnv
Proj_name = "HFSS_100G EML TOSA_0322"
Design_name = "Cut2"
Out_path = "D:/Work/COC_HHI_NEO/Result/0322/with_COC_chip_Res_Cap_0329_"
export_all_parameter = True
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oProject = oDesktop.SetActiveProject(Proj_name)
oDesign = oProject.SetActiveDesign(Design_name)
oModule = oDesign.GetModule("ReportSetup")
oModule.ExportToFile("TDR", Out_path+"TDR.csv",export_all_parameter)
oModule.ExportToFile("S11", Out_path+"S11.csv",export_all_parameter)
oModule.ExportToFile("S21", Out_path+"S21.csv",export_all_parameter)
oModule.ExportToFile("S22", Out_path+"S22.csv",export_all_parameter)
