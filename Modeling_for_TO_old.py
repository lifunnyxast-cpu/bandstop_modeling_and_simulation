"""
encode with UTF-8
PYTHON version 3.11.4 64-bit
CST Studio version 2022
"""

import time
import win32com.client
import os
import math
import numpy as np
import pd
import pandas as pdd
def quotes_to_list(text):
    """
    这是一个将跨行字符串分行转化为数组的函数
    一般我在将VBA语句大卸八块的时候会用这个函数作为辅助
    """
    lines = text.splitlines()
    quoted_lines = ["" + line.lstrip() + "" for line in lines]
    return quoted_lines


class StructureMacros:
    """
    CST 建模与设置的基类，封装了与历史树 (History Tree) 交互的底层逻辑
    """

    def __init__(self, handle):
        self.mws = handle
        self.Classname = self.__class__.__name__
        if self.mws is None:
            raise ValueError("MWS未进行初始化，请重新进行mws初始化")
        else:
            print(f"[{self.Classname}] MWS 句柄绑定成功{' ' * 10}", end="\r")

    def AddToHistoryWithList(self, Tag, Command):
        line_break = "\n"
        Command = line_break.join(Command)
        self.mws._FlagAsMethod("AddToHistory")
        self.mws.AddToHistory(Tag, Command)

    def AddToHistoryWithCommand(self, Tag, Command):
        self.mws._FlagAsMethod("AddToHistory")
        self.mws.AddToHistory(Tag, Command)


# ==============================================================================
# 2. 业务功能类定义
# ==============================================================================
#新建CST项目文件及初始化
class CSTInterface(StructureMacros):
    def __init__(self, label="New", project_name="topology_test"):
        self.cst = self.OpenCST()
        self.working_dir = os.getcwd()
        self.project_path = os.path.join(self.working_dir, f"{project_name}.cst")

        if label == "New":
            self.mws = self.NewProject()
            super().__init__(self.mws)

            self.custom_unit_initial()
            self.custom_background_initial()
            self.custom_FloquetPort_initial()
            self.custom_boundary_initial()
            self.custom_mesh_initial()
            self.custom_FDSolver_initial()
            self.save_as_project(self.project_path)

        elif label == "Open":
            self.mws = self.OpenProject(self.project_path)
            super().__init__(self.mws)

    def OpenCST(self):
        return win32com.client.dynamic.Dispatch("CSTStudio.Application")

    def NewProject(self):
        self.cst.NewMWS()
        return self.cst.Active3D()

    def OpenProject(self, ProjectName):
        self.cst.OpenFile(ProjectName)
        return self.cst.Active3D()

    def save_as_project(self, project_path):
        self.mws._FlagAsMethod("SaveAs")
        self.mws.SaveAs(project_path, "false")

    def quit_project_without_saving(self):
        """直接退出当前工程不保存，用于清理环境"""
        try:
            self.mws.Quit()
        except:
            pass

    def custom_unit_initial(self):
        initialsCommand = [
            "With Units", '.Geometry "mm"', '.Frequency "GHz"', '.Voltage "V"',
            '.Resistance "Ohm"', '.Inductance "H"', '.TemperatureUnit "Kelvin"',
            '.Time "ns"', '.Current "A"', '.Conductance "Siemens"', '.Capacitance "F"',
            "End With"
        ]
        self.AddToHistoryWithList("Unit Initial", initialsCommand)

    def custom_background_initial(self):
        sCommand = """
        'set the frequency range
        Solver.FrequencyRange "0", "20"

        Plot.DrawBox True
        With Background
             .Type "Normal"
             .Epsilon "1.0"
             .Mu "1.0"
             .Rho "1.204"
             .ThermalType "Normal"
             .ThermalConductivity "0.026"
              .SpecificHeat "1005", "J/K/kg"
             .XminSpace "0.0"
             .XmaxSpace "0.0"
             .YminSpace "0.0"
             .YmaxSpace "0.0"
             .ZminSpace "0.0"
             .ZmaxSpace "0.0"
        End With
        """
        self.AddToHistoryWithCommand("Background Initial", sCommand)

    def custom_FloquetPort_initial(self):
        vba = """
         With FloquetPort
             .Reset
             .SetDialogTheta "0"
             .SetDialogPhi "0"
             .SetSortCode "+beta/pw"
             .SetCustomizedListFlag "False"
             .Port "Zmin"
             .SetNumberOfModesConsidered "2"
             .Port "Zmax"
             .SetNumberOfModesConsidered "2"
        End With

        MakeSureParameterExists "theta", "0"
        SetParameterDescription "theta", "spherical angle of incident plane wave"
        MakeSureParameterExists "phi", "0"
        SetParameterDescription "phi", "spherical angle of incident plane wave"
         """
        self.AddToHistoryWithCommand("Clear Component", vba)

    def custom_boundary_initial(self):
        sCommand = """
        With Boundary
             .Xmin "unit cell"
             .Xmax "unit cell"
             .Ymin "unit cell"
             .Ymax "unit cell"
             .Zmin "open"
             .Zmax "open"
             .Xsymmetry "none"
             .Ysymmetry "none"
             .Zsymmetry "none"
             .XPeriodicShift "0.0"
             .YPeriodicShift "0.0"
             .ZPeriodicShift "0.0"
             .PeriodicUseConstantAngles "False"
             .SetPeriodicBoundaryAngles "theta", "phi"
             .SetPeriodicBoundaryAnglesDirection "inward"
             .UnitCellFitToBoundingBox "True"
             .UnitCellDs1 "0.0"
             .UnitCellDs2 "0.0"
             .UnitCellAngle "90.0"
        End With        
        """
        self.AddToHistoryWithCommand("Boundary Initial", sCommand)

    def custom_mesh_initial(self):
        sCommand = """
        With Mesh
             .MeshType "Tetrahedral"
        End With
        """
        self.AddToHistoryWithCommand("Mesh Initial", sCommand)

    def custom_FDSolver_initial(self):
        sCommand = """
        With FDSolver
             .Reset
             .Stimulation "List", "List"
             .ResetExcitationList
             .AddToExcitationList "Zmax", "TE(0,0);TM(0,0)"
             .LowFrequencyStabilization "False"
        End With
        ChangeSolverType("HF Frequency Domain")
        """
        self.AddToHistoryWithCommand("FDSolver Initial", sCommand)

    def StoreParameters(self, parameterspath, encodewith="utf-8"):
        # 首先将指定路径的文件一行一行的读取出来
        # 文件可以是CST直接导出的变量文件
        # 创建一个空字典来存储变量信息
        variables = {}

        # 打开文件并读取每一行
        with open(parameterspath, "r", encoding=encodewith) as f:
            for line in f:
                if line.strip() == "":#如果有空行，那么就跳过这个循环，也就跳过空行了
                    continue
                # 分割每一行并获取变量名和值
                parts = line.split("=")
                var_name = parts[0].strip()
                var_value, var_description = (
                    parts[1].split('"')[1],
                    parts[1].split('"')[3],
                )
                # 将变量名、值和描述存储在字典中
                variables[var_name] = {
                    "value": var_value,
                    "description": var_description,
                }
        self.ProcessingDictionary(variables)
        return variables

    def ProcessingDictionary(self, variables):
        sCommand = ""
        for key, value in variables.items():
            sCommand += """  
                MakeSureParameterExists("{0}", "{1}")
                SetParameterDescription  ( "{2}", "{3}" )
            """.format(
                key, value["value"], key, value["description"]
            )
        self.AddToHistoryWithCommand("Storeparameters", sCommand)


    # def run_simulation(self):
    #     print(">>> [SIM] 开始仿真 (F-Solver)...")
    #     # 启动 F-Solver
    #     vba = "FDSolver.Start"
    #     self.AddToHistoryWithCommand("FDSolver Start",vba)
    #
    #     # 启动后强制等待预热
    #     print(">>> [WAIT] ⏳ 仿真已启动，等待 5 秒...")
    #     time.sleep(5.0)
    #
    #     # 轮询等待结束
    #     self._wait_by_polling()

    def _wait_by_polling(self):
        """
        轮询监控 CST 状态，直到仿真结束
        """
        print(">>> [MONITOR] 开始监控仿真进度...")
        start_time = time.time()
        MAX_TIMEOUT = 3600  # 1小时超时

        # 监控标志：检查 1D Results 里是否生成了 S 参数
        # 注意：CST 运行中可能会锁定 Result Tree，所以用 try-catch

        while True:
            elapsed = int(time.time() - start_time)
            if elapsed > MAX_TIMEOUT:
                print("❌ [ERROR] 仿真超时！")
                break

            # 方法 A: 检查求解器是否还在运行
            # 这是一个比较通用的检测方法：尝试访问 Solver 状态
            # 但最简单的方法是检查是否有 S 参数生成

            # 这里我们简化处理：假设每 10 秒打印一次，
            # 实际上很难通过 COM 接口实时知道进度条，
            # 通常我们通过检查是否有错误或者是否生成了特定文件来判断。
            # 附件里的方法是 ExportPlotData 来检查，我们可以照搬。

            try:
                # 尝试导出一个特定的结果，如果成功且非空，说明算完了
                # 注意：Result Tree 路径在不同版本可能不同
                # CST 2022 通常是 "1D Results\S-Parameters\SZmax(1),Zmax(1)"
                res_path = r"1D Results\S-Parameters\SZmax(1),Zmax(1)"


                # freq, s_val = self.get_s_parameter(check_path, export_file="check_status.txt")

                # 尝试导出到临时文件
                temp_file_name = os.path.join(self.working_dir, "check_status.txt")
                if os.path.exists(temp_file_name): os.remove(temp_file_name)

                vba = f'SelectTreeItem "{res_path}"\nExportPlotData "{temp_file_name}"'
                ResultExtractor.extract_db_results(self.working_dir,out_name="check_status.txt")


                if os.path.exists(temp_file_name) and os.path.getsize(temp_file_name) > 100:
                    # 文件存在且有内容，说明算完了
                    print(f"✅ [DONE] 新数据生成！仿真结束 (耗时: {elapsed}s)")
                    return
            except:
                pass

            if elapsed % 10 == 0:
                print(f"   ... 已运行 {elapsed}s ...")

            time.sleep(2.0)

    def get_s_parameter(self, result_path, export_file="temp_s11.txt"):
        self.mws._FlagAsMethod("SelectTreeItem")
        try:
            self.mws.SelectTreeItem(result_path)
        except:
            return None, None

        abs_export_path = os.path.join(os.path.dirname(self.project_path), export_file)
        if os.path.exists(abs_export_path):
            try:
                os.remove(abs_export_path)
            except:
                return None, None

        self.mws._FlagAsMethod("ExportPlotData")
        try:
            self.mws.ExportPlotData(abs_export_path)
        except:
            return None, None

        time.sleep(0.2)

        try:
            data = np.loadtxt(abs_export_path, skiprows=2)
            if data.ndim == 1:
                freq = np.array([data[0]])
                s_db = np.array([data[1]])
            else:
                freq = data[:, 0]
                s_db = data[:, 1]
            return freq, s_db
        except:
            return None, None

    def export_project(self, save_name):
        cwd = os.getcwd()
        res_dir = os.path.join(cwd, "Results_CST")
        if not os.path.exists(res_dir): os.makedirs(res_dir)

        full_path = os.path.join(res_dir, f"{save_name}.cst")
        # full_path = full_path.replace("/", "\\")

        vba = f'SaveAs "{full_path}", "True"'
        self.run_vba_via_file(vba)
        print(f">>> [SAVE] 工程已保存: {save_name}")

#模型构建及环境变量更新
class SimulationBuilder(StructureMacros):
    def __init__(self, cst_instance):
        super().__init__(cst_instance.mws)
        self.cst_interface = cst_instance

    def update_environment(self, pp, period,sym_type):
        # vba_base = """
        #        SetParameterDescription "theta", "spherical angle of incident plane wave"
        #        """
        # self.AddToHistoryWithCommand("Init Global Params", vba_base)
        vba_env = f"""
        ' 设置背景为真空，Z方向留出计算空间
        With Background
             .ResetBackground
             .XminSpace "0.0"
             .XmaxSpace "0.0"
             .YminSpace "0.0"
             .YmaxSpace "0.0"
             .ZminSpace "{pp}"
             .ZmaxSpace "{pp}"
             .ApplyInAllDirections "False"
        End With
        """
        self.AddToHistoryWithCommand("define Background", vba_env)
        #.SetPeriodicBoundaryAngles "theta", "phi"这句声明很关键，将角度与端口联系起来，否则扫角都是0°结果
        match sym_type:
            case (1 | 3):
                vba_env=f"""
                ' 设置 X,Y 方向为 unit cell (周期边界)，Z方向为 open (吸收边界)
                With Boundary
                     .Xmin "unit cell"
                     .Xmax "unit cell"
                     .Ymin "unit cell"
                     .Ymax "unit cell"
                     .Zmin "open"
                     .Zmax "open"
                     .Xsymmetry "none"
                     .Ysymmetry "none"
                     .Zsymmetry "none"
                     .ApplyInAllDirections "False"
                     .XPeriodicShift "0.0"
                     .YPeriodicShift "0.0"
                     .ZPeriodicShift "0.0"
                     .PeriodicUseConstantAngles "False"
                     .SetPeriodicBoundaryAngles "theta", "phi"
                     .SetPeriodicBoundaryAnglesDirection "inward"
                     .UnitCellFitToBoundingBox "False"
                     .UnitCellDs1 "{period*math.sqrt(3)}"
                     .UnitCellDs2 "{period*math.sqrt(3)}"
                     .UnitCellAngle "60.0"
                    End With
                    """
            case _:
                vba_env=f"""
                ' 设置 X,Y 方向为 unit cell (周期边界)，Z方向为 open (吸收边界)
                With Boundary
                     .Xmin "unit cell"
                     .Xmax "unit cell"
                     .Ymin "unit cell"
                     .Ymax "unit cell"
                     .Zmin "open"
                     .Zmax "open"
                     .Xsymmetry "none"
                     .Ysymmetry "none"
                     .Zsymmetry "none"
                     .ApplyInAllDirections "False"
                     .XPeriodicShift "0.0"
                     .YPeriodicShift "0.0"
                     .ZPeriodicShift "0.0"
                     .PeriodicUseConstantAngles "False"
                     .SetPeriodicBoundaryAngles "theta", "phi"
                     .SetPeriodicBoundaryAnglesDirection "inward"
                     .UnitCellFitToBoundingBox "False"
                     .UnitCellDs1 "{period}"
                     .UnitCellDs2 "{period}"
                     .UnitCellAngle "90.0"
                End With
                """
        self.AddToHistoryWithCommand("define Boundary", vba_env)
        vba_env="""
        With Mesh 
             .MeshType "Tetrahedral" 
             .SetCreator "High Frequency"
        End With 
        With MeshSettings 
             .SetMeshType "Tet" 
             .Set "Version", 1%
             'MAX CELL - WAVELENGTH REFINEMENT 
             .Set "StepsPerWaveNear", "4" 
             .Set "StepsPerWaveFar", "4" 
             .Set "PhaseErrorNear", "0.02" 
             .Set "PhaseErrorFar", "0.02" 
             .Set "CellsPerWavelengthPolicy", "automatic" 
             'MAX CELL - GEOMETRY REFINEMENT 
             .Set "StepsPerBoxNear", "15" 
             .Set "StepsPerBoxFar", "10" 
             .Set "ModelBoxDescrNear", "maxedge" 
             .Set "ModelBoxDescrFar", "maxedge" 
             'MIN CELL 
             .Set "UseRatioLimit", "0" 
             .Set "RatioLimit", "100" 
             .Set "MinStep", "0" 
             'MESHING METHOD 
             .SetMeshType "Unstr" 
             .Set "Method", "0" 
        End With 
        With MeshSettings 
             .SetMeshType "Tet" 
             .Set "CurvatureOrder", "1" 
             .Set "CurvatureOrderPolicy", "automatic" 
             .Set "CurvRefinementControl", "NormalTolerance" 
             .Set "NormalTolerance", "22.5" 
             .Set "SrfMeshGradation", "1.5" 
             .Set "SrfMeshOptimization", "1" 
        End With 
        With MeshSettings 
             .SetMeshType "Unstr" 
             .Set "UseMaterials",  "1" 
             .Set "MoveMesh", "0" 
        End With 
        With MeshSettings 
             .SetMeshType "All" 
             .Set "AutomaticEdgeRefinement",  "0" 
        End With 
        With MeshSettings 
             .SetMeshType "Tet" 
             .Set "UseAnisoCurveRefinement", "1" 
             .Set "UseSameSrfAndVolMeshGradation", "1" 
             .Set "VolMeshGradation", "1.5" 
             .Set "VolMeshOptimization", "1" 
        End With 
        With MeshSettings 
             .SetMeshType "Unstr" 
             .Set "SmallFeatureSize", "0" 
             .Set "CoincidenceTolerance", "1e-06" 
             .Set "SelfIntersectionCheck", "1" 
             .Set "OptimizeForPlanarStructures", "0" 
        End With 
        With Mesh 
             .SetParallelMesherMode "Tet", "maximum" 
             .SetMaxParallelMesherThreads "Tet", "1" 
        End With     
        """
        self.AddToHistoryWithCommand("update mesh",vba_env)
        vba_env="""
        Mesh.SetCreator "High Frequency" 

        With FDSolver
             .Reset 
             .SetMethod "Tetrahedral", "General purpose" 
             .OrderTet "Second" 
             .OrderSrf "First" 
             .Stimulation "Zmax", "All" 
             .ResetExcitationList 
             .AddToExcitationList "Zmax", "TE(0,0);TM(0,0)" 
             .AutoNormImpedance "False" 
             .NormingImpedance "50" 
             .ModesOnly "False" 
             .ConsiderPortLossesTet "True" 
             .SetShieldAllPorts "False" 
             .AccuracyHex "1e-6" 
             .AccuracyTet "1e-4" 
             .AccuracySrf "1e-3" 
             .LimitIterations "False" 
             .MaxIterations "0" 
             .SetCalcBlockExcitationsInParallel "True", "True", "" 
             .StoreAllResults "False" 
             .StoreResultsInCache "False" 
             .UseHelmholtzEquation "True" 
             .LowFrequencyStabilization "False" 
             .Type "Auto" 
             .MeshAdaptionHex "False" 
             .MeshAdaptionTet "True" 
             .AcceleratedRestart "True" 
             .FreqDistAdaptMode "Distributed" 
             .NewIterativeSolver "True" 
             .TDCompatibleMaterials "False" 
             .ExtrudeOpenBC "False" 
             .SetOpenBCTypeHex "Default" 
             .SetOpenBCTypeTet "Default" 
             .AddMonitorSamples "True" 
             .CalcPowerLoss "True" 
             .CalcPowerLossPerComponent "False" 
             .StoreSolutionCoefficients "True" 
             .UseDoublePrecision "False" 
             .UseDoublePrecision_ML "True" 
             .MixedOrderSrf "False" 
             .MixedOrderTet "False" 
             .PreconditionerAccuracyIntEq "0.15" 
             .MLFMMAccuracy "Default" 
             .MinMLFMMBoxSize "0.3" 
             .UseCFIEForCPECIntEq "True" 
             .UseEnhancedCFIE2 "True" 
             .UseFastRCSSweepIntEq "false" 
             .UseSensitivityAnalysis "False" 
             .UseEnhancedNFSImprint "False" 
             .RemoveAllStopCriteria "Hex"
             .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
             .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
             .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
             .RemoveAllStopCriteria "Tet"
             .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
             .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
             .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
             .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
             .RemoveAllStopCriteria "Srf"
             .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
             .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
             .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
             .SweepMinimumSamples "3" 
             .SetNumberOfResultDataSamples "1001" 
             .SetResultDataSamplingMode "Automatic" 
             .SweepWeightEvanescent "1.0" 
             .AccuracyROM "1e-4" 
             .AddSampleInterval "", "", "1", "Automatic", "True" 
             .AddSampleInterval "", "", "", "Automatic", "False" 
             .MPIParallelization "False"
             .UseDistributedComputing "False"
             .NetworkComputingStrategy "RunRemote"
             .NetworkComputingJobCount "3"
             .UseParallelization "True"
             .MaxCPUs "1024"
             .MaximumNumberOfCPUDevices "2"
        End With
        
        With IESolver
             .Reset 
             .UseFastFrequencySweep "True" 
             .UseIEGroundPlane "False" 
             .SetRealGroundMaterialName "" 
             .CalcFarFieldInRealGround "False" 
             .RealGroundModelType "Auto" 
             .PreconditionerType "Auto" 
             .ExtendThinWireModelByWireNubs "False" 
             .ExtraPreconditioning "False" 
        End With
        
        With IESolver
             .SetFMMFFCalcStopLevel "0" 
             .SetFMMFFCalcNumInterpPoints "6" 
             .UseFMMFarfieldCalc "True" 
             .SetCFIEAlpha "0.500000" 
             .LowFrequencyStabilization "False" 
             .LowFrequencyStabilizationML "True" 
             .Multilayer "False" 
             .SetiMoMACC_I "0.0001" 
             .SetiMoMACC_M "0.0001" 
             .DeembedExternalPorts "True" 
             .SetOpenBC_XY "True" 
             .OldRCSSweepDefintion "False" 
             .SetRCSOptimizationProperties "True", "100", "0.00001" 
             .SetAccuracySetting "Custom" 
             .CalculateSParaforFieldsources "True" 
             .ModeTrackingCMA "True" 
             .NumberOfModesCMA "3" 
             .StartFrequencyCMA "-1.0" 
             .SetAccuracySettingCMA "Default" 
             .FrequencySamplesCMA "0" 
             .SetMemSettingCMA "Auto" 
             .CalculateModalWeightingCoefficientsCMA "True" 
             .DetectThinDielectrics "True" 
        End With

        """
        self.AddToHistoryWithCommand("update fd solver", vba_env)
    def update_environment_for_localmesh(self, name, pp, period, sym_type):
        vba_env = f"""
              ' 设置背景为真空，Z方向留出计算空间
              With Background
                   .ResetBackground
                   .XminSpace "0.0"
                   .XmaxSpace "0.0"
                   .YminSpace "0.0"
                   .YmaxSpace "0.0"
                   .ZminSpace "{pp}"
                   .ZmaxSpace "{pp}"
                   .ApplyInAllDirections "False"
              End With
              """
        self.AddToHistoryWithCommand("define Background", vba_env)
        # .SetPeriodicBoundaryAngles "theta", "phi"这句声明很关键，将角度与端口联系起来，否则扫角都是0°结果
        if sym_type in (1, 3):
            vba_env = f"""
                  ' 设置 X,Y 方向为 unit cell (周期边界)，Z方向为 open (吸收边界)
                  With Boundary
                       .Xmin "unit cell"
                       .Xmax "unit cell"
                       .Ymin "unit cell"
                       .Ymax "unit cell"
                       .Zmin "open"
                       .Zmax "open"
                       .Xsymmetry "none"
                       .Ysymmetry "none"
                       .Zsymmetry "none"
                       .ApplyInAllDirections "False"
                       .XPeriodicShift "0.0"
                       .YPeriodicShift "0.0"
                       .ZPeriodicShift "0.0"
                       .PeriodicUseConstantAngles "False"
                       .SetPeriodicBoundaryAngles "theta", "phi"
                       .SetPeriodicBoundaryAnglesDirection "inward"
                       .UnitCellFitToBoundingBox "False"
                       .UnitCellDs1 "{period*math.sqrt(3)}"
                       .UnitCellDs2 "{period*math.sqrt(3)}"
                       .UnitCellAngle "60.0"
                      End With
                      """
        else:
            vba_env = f"""
                  ' 设置 X,Y 方向为 unit cell (周期边界)，Z方向为 open (吸收边界)
                  With Boundary
                       .Xmin "unit cell"
                       .Xmax "unit cell"
                       .Ymin "unit cell"
                       .Ymax "unit cell"
                       .Zmin "open"
                       .Zmax "open"
                       .Xsymmetry "none"
                       .Ysymmetry "none"
                       .Zsymmetry "none"
                       .ApplyInAllDirections "False"
                       .XPeriodicShift "0.0"
                       .YPeriodicShift "0.0"
                       .ZPeriodicShift "0.0"
                       .PeriodicUseConstantAngles "False"
                       .SetPeriodicBoundaryAngles "theta", "phi"
                       .SetPeriodicBoundaryAnglesDirection "inward"
                       .UnitCellFitToBoundingBox "False"
                       .UnitCellDs1 "{period}"
                       .UnitCellDs2 "{period}"
                       .UnitCellAngle "90.0"
                  End With
                  """
        self.AddToHistoryWithCommand("define Boundary", vba_env)
        vba_env = """
              With Mesh 
                   .MeshType "Tetrahedral" 
                   .SetCreator "High Frequency"
              End With 
              With MeshSettings 
                   .SetMeshType "Tet" 
                   .Set "Version", 1%
                   'MAX CELL - WAVELENGTH REFINEMENT 
                   .Set "StepsPerWaveNear", "4" 
                   .Set "StepsPerWaveFar", "4" 
                   .Set "PhaseErrorNear", "0.02" 
                   .Set "PhaseErrorFar", "0.02" 
                   .Set "CellsPerWavelengthPolicy", "automatic" 
                   'MAX CELL - GEOMETRY REFINEMENT 
                   .Set "StepsPerBoxNear", "15" 
                   .Set "StepsPerBoxFar", "10" 
                   .Set "ModelBoxDescrNear", "maxedge" 
                   .Set "ModelBoxDescrFar", "maxedge" 
                   'MIN CELL 
                   .Set "UseRatioLimit", "0" 
                   .Set "RatioLimit", "100" 
                   .Set "MinStep", "0" 
                   'MESHING METHOD 
                   .SetMeshType "Unstr" 
                   .Set "Method", "0" 
              End With 
              With MeshSettings 
                   .SetMeshType "Tet" 
                   .Set "CurvatureOrder", "1" 
                   .Set "CurvatureOrderPolicy", "automatic" 
                   .Set "CurvRefinementControl", "NormalTolerance" 
                   .Set "NormalTolerance", "22.5" 
                   .Set "SrfMeshGradation", "1.5" 
                   .Set "SrfMeshOptimization", "1" 
              End With 
              With MeshSettings 
                   .SetMeshType "Unstr" 
                   .Set "UseMaterials",  "1" 
                   .Set "MoveMesh", "0" 
              End With 
              With MeshSettings 
                   .SetMeshType "All" 
                   .Set "AutomaticEdgeRefinement",  "0" 
              End With 
              With MeshSettings 
                   .SetMeshType "Tet" 
                   .Set "UseAnisoCurveRefinement", "1" 
                   .Set "UseSameSrfAndVolMeshGradation", "1" 
                   .Set "VolMeshGradation", "1.5" 
                   .Set "VolMeshOptimization", "1" 
              End With 
              With MeshSettings 
                   .SetMeshType "Unstr" 
                   .Set "SmallFeatureSize", "0" 
                   .Set "CoincidenceTolerance", "1e-06" 
                   .Set "SelfIntersectionCheck", "1" 
                   .Set "OptimizeForPlanarStructures", "1" 
              End With 
              With Mesh 
                   .SetParallelMesherMode "Tet", "maximum" 
                   .SetMaxParallelMesherThreads "Tet", "1" 
              End With     
              """
        self.AddToHistoryWithCommand("update mesh", vba_env)

        vba_env = """
            Group.Add "meshgroup1", "mesh"
        """
        self.AddToHistoryWithCommand("creat group:mesh group1", vba_env)
        # 新增局部网格加密
        vba_env = """
        Group.Add "meshgroup1", "mesh"
            With MeshSettings
                 With .ItemMeshSettings ("group$meshgroup1")
                      .SetMeshType "Tet"
                      .Set "LayerStackup", "Automatic"
                      .Set "LocalAutomaticEdgeRefinement", "0"
                      .Set "LocalAutomaticEdgeRefinementOverwrite", 0
                      .Set "MaterialIndependent", 0
                      .Set "OctreeSizeFaces", "0"
                      .Set "PatchIndependent", 0
                      .Set "Size", "0.2"
                 End With
            End With
            """
        self.AddToHistoryWithCommand("set local mesh for：meshgroup1", vba_env)
        vba_env = f"""
                  Group.AddItem "solid$component1:{name}", "meshgroup1"
              """
        self.AddToHistoryWithCommand("add items to group：meshgroup1", vba_env)
        vba_env = """
              Mesh.SetCreator "High Frequency" 

              With FDSolver
                   .Reset 
                   .SetMethod "Tetrahedral", "General purpose" 
                   .OrderTet "Second" 
                   .OrderSrf "First" 
                   .Stimulation "Zmax", "All" 
                   .ResetExcitationList 
                   .AddToExcitationList "Zmax", "TE(0,0);TM(0,0)" 
                   .AutoNormImpedance "False" 
                   .NormingImpedance "50" 
                   .ModesOnly "False" 
                   .ConsiderPortLossesTet "True" 
                   .SetShieldAllPorts "False" 
                   .AccuracyHex "1e-6" 
                   .AccuracyTet "1e-4" 
                   .AccuracySrf "1e-3" 
                   .LimitIterations "False" 
                   .MaxIterations "0" 
                   .SetCalcBlockExcitationsInParallel "True", "True", "" 
                   .StoreAllResults "False" 
                   .StoreResultsInCache "False" 
                   .UseHelmholtzEquation "True" 
                   .LowFrequencyStabilization "False" 
                   .Type "Auto" 
                   .MeshAdaptionHex "False" 
                   .MeshAdaptionTet "False" 
                   .AcceleratedRestart "True" 
                   .FreqDistAdaptMode "Distributed" 
                   .NewIterativeSolver "True" 
                   .TDCompatibleMaterials "False" 
                   .ExtrudeOpenBC "False" 
                   .SetOpenBCTypeHex "Default" 
                   .SetOpenBCTypeTet "Default" 
                   .AddMonitorSamples "True" 
                   .CalcPowerLoss "True" 
                   .CalcPowerLossPerComponent "False" 
                   .StoreSolutionCoefficients "True" 
                   .UseDoublePrecision "False" 
                   .UseDoublePrecision_ML "True" 
                   .MixedOrderSrf "False" 
                   .MixedOrderTet "False" 
                   .PreconditionerAccuracyIntEq "0.15" 
                   .MLFMMAccuracy "Default" 
                   .MinMLFMMBoxSize "0.3" 
                   .UseCFIEForCPECIntEq "True" 
                   .UseEnhancedCFIE2 "True" 
                   .UseFastRCSSweepIntEq "false" 
                   .UseSensitivityAnalysis "False" 
                   .UseEnhancedNFSImprint "False" 
                   .RemoveAllStopCriteria "Hex"
                   .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
                   .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
                   .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
                   .RemoveAllStopCriteria "Tet"
                   .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
                   .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
                   .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
                   .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
                   .RemoveAllStopCriteria "Srf"
                   .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
                   .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
                   .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
                   .SweepMinimumSamples "3" 
                   .SetNumberOfResultDataSamples "1001" 
                   .SetResultDataSamplingMode "Automatic" 
                   .SweepWeightEvanescent "1.0" 
                   .AccuracyROM "1e-4" 
                   .AddSampleInterval "", "", "1", "Automatic", "True" 
                   .AddSampleInterval "", "", "", "Automatic", "False" 
                   .MPIParallelization "False"
                   .UseDistributedComputing "False"
                   .NetworkComputingStrategy "RunRemote"
                   .NetworkComputingJobCount "3"
                   .UseParallelization "True"
                   .MaxCPUs "1024"
                   .MaximumNumberOfCPUDevices "2"
              End With

              With IESolver
                   .Reset 
                   .UseFastFrequencySweep "True" 
                   .UseIEGroundPlane "False" 
                   .SetRealGroundMaterialName "" 
                   .CalcFarFieldInRealGround "False" 
                   .RealGroundModelType "Auto" 
                   .PreconditionerType "Auto" 
                   .ExtendThinWireModelByWireNubs "False" 
                   .ExtraPreconditioning "False" 
              End With

              With IESolver
                   .SetFMMFFCalcStopLevel "0" 
                   .SetFMMFFCalcNumInterpPoints "6" 
                   .UseFMMFarfieldCalc "True" 
                   .SetCFIEAlpha "0.500000" 
                   .LowFrequencyStabilization "False" 
                   .LowFrequencyStabilizationML "True" 
                   .Multilayer "False" 
                   .SetiMoMACC_I "0.0001" 
                   .SetiMoMACC_M "0.0001" 
                   .DeembedExternalPorts "True" 
                   .SetOpenBC_XY "True" 
                   .OldRCSSweepDefintion "False" 
                   .SetRCSOptimizationProperties "True", "100", "0.00001" 
                   .SetAccuracySetting "Custom" 
                   .CalculateSParaforFieldsources "True" 
                   .ModeTrackingCMA "True" 
                   .NumberOfModesCMA "3" 
                   .StartFrequencyCMA "-1.0" 
                   .SetAccuracySettingCMA "Default" 
                   .FrequencySamplesCMA "0" 
                   .SetMemSettingCMA "Auto" 
                   .CalculateModalWeightingCoefficientsCMA "True" 
                   .DetectThinDielectrics "True" 
              End With

              """
        self.AddToHistoryWithCommand("update fd solver", vba_env)
    def create_materials(self):
        vba_mat = """
       With Material
             .Reset
             .Name "RB_Substrate"
             .Folder ""
             .Rho "0.0"
             .ThermalType "Normal"
             .ThermalConductivity "0"
             .SpecificHeat "0", "J/K/kg"
             .DynamicViscosity "0"
             .Emissivity "0"
             .MetabolicRate "0.0"
             .VoxelConvection "0.0"
             .BloodFlow "0"
             .MechanicsType "Unused"
             .IntrinsicCarrierDensity "0"
             .FrqType "all"
             .Type "Normal"
             .MaterialUnit "Frequency", "GHz"
             .MaterialUnit "Geometry", "mm"
             .MaterialUnit "Time", "ns"
             .MaterialUnit "Temperature", "Kelvin"
             .Epsilon "3.2"
             .Mu "1"
             .Sigma "0.0"
             .TanD "0.002"
             .TanDFreq "10"
             .TanDGiven "True"
             .TanDModel "ConstTanD"
             .SetConstTanDStrategyEps "AutomaticOrder"
             .ConstTanDModelOrderEps "3"
             .DjordjevicSarkarUpperFreqEps "0"
             .SetElParametricConductivity "False"
             .ReferenceCoordSystem "Global"
             .CoordSystemType "Cartesian"
             .SigmaM "0"
             .TanDM "0.0"
             .TanDMFreq "0.0"
             .TanDMGiven "False"
             .TanDMModel "ConstTanD"
             .SetConstTanDStrategyMu "AutomaticOrder"
             .ConstTanDModelOrderMu "3"
             .DjordjevicSarkarUpperFreqMu "0"
             .SetMagParametricConductivity "False"
             .DispModelEps "None"
             .DispModelMu "None"
             .DispersiveFittingSchemeEps "Nth Order"
             .MaximalOrderNthModelFitEps "10"
             .ErrorLimitNthModelFitEps "0.1"
             .UseOnlyDataInSimFreqRangeNthModelEps "False"
             .DispersiveFittingSchemeMu "Nth Order"
             .MaximalOrderNthModelFitMu "10"
             .ErrorLimitNthModelFitMu "0.1"
             .UseOnlyDataInSimFreqRangeNthModelMu "False"
             .UseGeneralDispersionEps "False"
             .UseGeneralDispersionMu "False"
             .NLAnisotropy "False"
             .NLAStackingFactor "1"
             .NLADirectionX "1"
             .NLADirectionY "0"
             .NLADirectionZ "0"
             .Colour "0", "1", "1"
             .Wireframe "False"
             .Reflection "False"
             .Allowoutline "True"
             .Transparentoutline "False"
             .Transparency "0"
             .Create
        End With
        With Material
             .Reset
             .Name "Copper (annealed)"
             .Folder ""
            .FrqType "static"
            .Type "Normal"
            .SetMaterialUnit "Hz", "mm"
            .Epsilon "1"
            .Mu "1.0"
            .Kappa "5.8e+007"
            .TanD "0.0"
            .TanDFreq "0.0"
            .TanDGiven "False"
            .TanDModel "ConstTanD"
            .KappaM "0"
            .TanDM "0.0"
            .TanDMFreq "0.0"
            .TanDMGiven "False"
            .TanDMModel "ConstTanD"
            .DispModelEps "None"
            .DispModelMu "None"
            .DispersiveFittingSchemeEps "Nth Order"
            .DispersiveFittingSchemeMu "Nth Order"
            .UseGeneralDispersionEps "False"
            .UseGeneralDispersionMu "False"
            .FrqType "all"
            .Type "Lossy metal"
            .SetMaterialUnit "GHz", "mm"
            .Mu "1.0"
            .Kappa "5.8e+007"
            .Rho "8930.0"
            .ThermalType "Normal"
            .ThermalConductivity "401.0"
            .SpecificHeat "390", "J/K/kg"
            .MetabolicRate "0"
            .BloodFlow "0"
            .VoxelConvection "0"
            .MechanicsType "Isotropic"
            .YoungsModulus "120"
            .PoissonsRatio "0.33"
            .ThermalExpansionRate "17"
            .Colour "1", "1", "0"
            .Wireframe "False"
            .Reflection "False"
            .Allowoutline "True"
            .Transparentoutline "False"
            .Transparency "0"
            .Create
        End With
        """
        self.AddToHistoryWithCommand("Create Materials", vba_mat)

    def build_geometry(self, period,sym_type,line_width,name, points):
        h_copper = 0.018
        h_sub = 0.025
        vba_sub=None
        match sym_type:
            case (1 | 3):
                vba_sub = f"""
                With Cylinder 
                 .Reset 
                 .Name "Substrate" 
                 .Component "component1" 
                 .Material "RB_Substrate" 
                 .OuterRadius "{period}" 
                 .InnerRadius "0" 
                 .Axis "z" 
                 .Zrange "{-h_sub}", "0" 
                 .Xcenter "0" 
                 .Ycenter "0" 
                 .Segments "6" 
                 .Create 
                End With
                """
            case _:
                vba_sub = f"""  
                With Brick
                    .Reset
                    .Name "Substrate"
                    .Component "component1"
                    .Material "RB_Substrate"
                    .Xrange "{-period / 2}", "{period / 2}"
                    .Yrange "{-period / 2}", "{period / 2}"
                    .Zrange "0", "{-h_sub}"
                    .Create
                End With
                """
        self.AddToHistoryWithCommand("Build Substrate", vba_sub)

        vba_sub = f"""
        Curve.NewCurve "curve1"
        With Polygon
            .Reset
            .Name "{name}"
            .Curve "curve1"
        """
        vba_sub += f'    .Point "{points[0][0]:.4f}", "{points[0][1]:.4f}"\n'
        for pt in points[1:]:
            vba_sub += f'    .RLine "{pt[0]:.4f}", "{pt[1]:.4f}"\n'
        vba_sub += """    .Create
        End With
        """
        self.AddToHistoryWithCommand("Draw Curve", vba_sub)

        vba_sub = f"""
        With TraceFromCurve
            .Reset
            .Name "{name}"
            .Component "component1"
            .Curve "curve1"
            .Material "Copper (annealed)"
            .Thickness "{h_copper}"
            .Width "{line_width}"
            .RoundEnd "True"
            .Create
        End With
        Curve.DeleteCurve "curve1"
        """
        self.AddToHistoryWithCommand("Create Trace", vba_sub)

        copies = 0
        angle = 0
        if sym_type == 1:
            copies, angle = 6, 60
        elif sym_type == 2:
            copies, angle = 4, 90
        elif sym_type == 3:
            copies, angle = 3, 120
        elif sym_type == 4:
            copies, angle = 2, 180

        if copies > 1:
            vba_rot = f"""
            With Transform
                .Reset
                .Name "component1:{name}"
                .Origin "Free"
                .Center "0", "0", "0"
                .Angle "0", "0", "{angle}"
                .MultipleObjects "True"
                .GroupObjects "False"
                .Repetitions "{copies - 1}"
                .MultipleSelection "False"
                .Destination ""
                .Material ""
                .AutoDestination "True"
                .Transform "Shape", "Rotate"
            End With
            """
            self.AddToHistoryWithCommand("Symmetry Rotate", vba_rot)

    def boolean_and_shapes(self, row):
        sym_type = int(row['Sym_Type'])
        name = str(row['Source_ID']).split('.')[0]
        copies = {1: 5, 2: 3, 3: 2, 4: 1}.get(sym_type, 0)

        for n in range(copies):
            vba_boolean = f'Solid.Add "component1:{name}", "component1:{name}_{n + 1}"\n'
            self.AddToHistoryWithCommand(f"boolean_and_shapes_{n}", vba_boolean)

    def transform_rotate(self):
        vba_rot="""
        With Transform 
             .Reset 
             .Name "component1" 
             .Origin "Free" 
             .Center "0", "0", "0" 
             .Angle "0", "0", "30" 
             .MultipleObjects "False" 
             .GroupObjects "False" 
             .Repetitions "1" 
             .MultipleSelection "False" 
             .AutoDestination "True" 
             .Transform "Shape", "Rotate" 
        End With
        """
        self.AddToHistoryWithCommand("Transform ：Rotate", vba_rot)

    def setup_floquet_ports(self, pp):
        """改进的Floquet端口配置：确保扫描角与参数扫描同步"""
        vba = f"""

        With FloquetPort
            .Reset
            ' 关键改动：使用参数变量 "theta" 作为扫描角
            .SetDialogTheta "theta"
            .SetDialogPhi "0"
            ' 确保极化与扫描角无关
            .SetPolarizationIndependentOfScanAnglePhi "0.0", "False"
            .SetSortCode "+beta/pw"
            .SetCustomizedListFlag "False"
            ' 下端口配置
            .Port "Zmin"
            .SetNumberOfModesConsidered "2"
            .SetDistanceToReferencePlane "{-pp}"
            .SetUseCircularPolarization "False"
            ' 上端口配置
            .Port "Zmax"
            .SetNumberOfModesConsidered "2"
            .SetDistanceToReferencePlane "{-pp}"
            .SetUseCircularPolarization "False"
        End With
        """
        self.AddToHistoryWithCommand("Setup Floquet Ports", vba)

    def clear_model(self):
        """
        按照 [删结果 -> 判存 -> 删模型 -> 刷新] 逻辑清理当前工程
        """
        print("    >> 执行模型环境清理...")

        # 1. 直接调用 COM 方法删除所有结果，彻底消除后续改动时的弹窗隐患
        try:
            self.mws._FlagAsMethod("DeleteResults")
            self.mws.DeleteResults()
        except Exception:
            # 如果是刚新建的空工程，或者上一次仿真失败没有结果，
            # DeleteResults 可能会抛出异常，直接 pass 忽略即可
            pass

            # 2. 将“判断是否存在”交由 CST VBA 引擎内部处理（瞬间完成，绝不报错）
        vba_delete_comp = """
        ' 检查 component1 是否存在，存在则删除，不存在则直接跳过
        If Component.DoesExist("component1") Then
            Component.Delete "component1"
        End If
        """
        self.AddToHistoryWithCommand("Clear Component", vba_delete_comp)

        # 3. 刷新模型 (更新视图、历史树和内部参数状态)
        try:
            # 尝试你提到的短指令
            self.mws.Rebuild()
        except Exception:
            # 兼容性补充：在大多数 CST COM 接口版本中，
            # 标准的强制全局刷新命令是 RebuildOnParametricChange
            self.mws.RebuildOnParametricChange()

        print("    >> 清理完成，环境已就绪。")

    def build(self, row, points):
        """完整的建模流程"""
        period = round(float(row['Var_Radius_mm']),2)
        pp = math.ceil(max(period * 0.5, 1.5))#math.ceil,向上取整
        sym_type = int(row['Sym_Type'])
        line_width = round(float(row['Var_Width_mm']),2)
        name = str(row['Source_ID']).split('.')[0]
        
        self.clear_model()
        self.create_materials()
        self.build_geometry(period,sym_type,line_width,name, points)
        self.boolean_and_shapes(row)
        if sym_type in (1, 3):
            self.transform_rotate()
            print(f"sym_type={sym_type},执行旋转30°旋转")
        else:
            print(f"sym_type={sym_type},不执行旋转")
        self.update_environment_for_localmesh(name,pp, period, sym_type)
        # self.update_environment(pp, period,sym_type)
        self.setup_floquet_ports(pp)

        # 6. 【终极防线】建模结果验证：检查拓扑实体是否存在
        # ---------------------------------------------------------
        print(f"    >> 正在验证拓扑实体 '{name}' 是否成功生成...")
        try:
            self.mws._FlagAsMethod("SelectTreeItem")
            # 拼接该实体在 CST 导航树中的标准路径
            tree_path = f"Components\\component1\\{name}"

            # SelectTreeItem 会返回 True (存在并成功选中) 或 False (未找到)
            is_exist = self.mws.SelectTreeItem(tree_path)

            if is_exist:
                print(f"    ✅ 建模成功：实体 '{name}' 确认存在。")
                return True
            else:
                print(f"    ❌ 建模失败：未生成 '{name}' (可能原因：路径自交、点重合或拉伸失败)。跳过此模型。")
                return False
        except Exception as e:
            # 如果因为版本兼容性导致接口调用报错，我们默认放行，防止阻塞自动化流程
            print(f"    [警告] 实体验证接口调用异常: {e}。默认放行。")
            return True


#扫参器设置
class SweepManager(StructureMacros):
    def __init__(self, cst_instance):
        super().__init__(cst_instance.mws)
        self.cst_interface = cst_instance

    def _wait_by_polling(self):
        print(">>> [MONITOR] 启用轮询监控...")
        start_time = time.time()

        check_path = r"1D Results\S-Parameters\SZmax(1),Zmax(1)"
        MAX_TIMEOUT = 1000  # 超时设置

        while True:
            # -------------------------
            # 超时判断
            # -------------------------
            if time.time() - start_time > MAX_TIMEOUT:
                print("❌ [ERROR] 仿真超时！")
                return False

            # 检测 CST 报错
            try:
                n_err = self.mws.GetNumberOfSolverErrorMessages() + self.mws.GetNumberOfAnalysisErrorMessages()
                if n_err > 0:
                    print(f"❌ [CST Error] 求解器报错！错误数量: {n_err}")
                    try:
                        self.mws.FDSolver.Quit()
                    except:
                        pass
                    return
            except:
                pass

            # -------------------------
            # 检查结果是否生成
            # -------------------------
            try:
                result_tree = self.mws.ResultTree
                exists = result_tree.DoesTreeItemExist(check_path)
                if exists:
                    elapsed = int(time.time() - start_time)
                    # 安全门槛 10s
                    if elapsed < 6:
                        print(f"⚠️ 数据生成过快 ({elapsed}s < 6s)，疑似旧缓存，继续等待...")
                        time.sleep(3.0)
                        continue

                    print(f"✅ 新数据生成！仿真结束 (耗时: {elapsed}s)")
                    return True
            except:
                pass
            # -------------------------
            # 每 间隔打印一次状态
            # -------------------------
            elapsed = int(time.time() - start_time)
            if elapsed % 10 == 0:
                print(f">>> 计算中... (已耗时 {elapsed}s)")

            time.sleep(2.0)

    # def get_s_parameter(self, result_path, export_file="temp_s11.txt"):
    #     self.mws._FlagAsMethod("SelectTreeItem")
    #     try:
    #         self.mws.SelectTreeItem(result_path)
    #     except:
    #         return None, None
    #
    #     abs_export_path = os.path.join(os.path.dirname(self.cst_path), export_file)
    #     if os.path.exists(abs_export_path):
    #         try:
    #             os.remove(abs_export_path)
    #         except:
    #             return None, None
    #
    #     self.mws._FlagAsMethod("ExportPlotData")
    #     try:
    #         self.mws.ExportPlotData(abs_export_path)
    #     except:
    #         return None, None
    #
    #     time.sleep(0.2)
    #
    #     try:
    #         data = np.loadtxt(abs_export_path, skiprows=2)
    #         if data.ndim == 1:
    #             freq = np.array([data[0]])
    #             s_db = np.array([data[1]])
    #         else:
    #             freq = data[:, 0]
    #             s_db = data[:, 1]
    #         return freq, s_db
    #     except:
    #         return None, None



    def run_sweep_task(self):
        # 1. 参数扫描配置：theta 从 0 到 80 度，步长 5 度
        vba_macro_content = """
                With ParameterSweep
                    .SetSimulationType ("Frequency")
                    .DeleteAllSequences
                    .AddSequence ("AngleScan")
                    .AddParameter_Samples "AngleScan", "theta", 0, 80, 5, False
                End With
                    """
        self.AddToHistoryWithCommand("Setup ParameterSweep", vba_macro_content)

        # 3. 执行仿真并等待完成
        print("    >> 开始执行 Parameter Sweep ... (请在 CST 界面查看进度)")
        try:
            # 直接调用 Start 属性执行仿真
            self.mws.ParameterSweep.Start
            # 启动后强制等待 8 秒
            # print(">>> [WAIT] ⏳ 仿真已启动，等待 8 秒预热...")
            # time.sleep(8.0)

            return self._wait_by_polling()#成功会返回True,失败会返回false
        except Exception as e:
            print(f"    >> [Error] Parameter Sweep 运行失败: {e}")
            return False

class ResultExtractor(StructureMacros):
    def __init__(self, cst_instance):#接收了 cst 实例，并将 CST 句柄保存到了 self.mws 中
        super().__init__(cst_instance.mws)
        self.cst_interface = cst_instance

    # def extract_db_results(self, out_path):
    #     """改进的结果提取：增强稳健性和错误处理"""
    #     modes_to_extract = [
    #         "SZmax(1),Zmax(1)", "SZmin(1),Zmax(1)",
    #         "SZmax(2),Zmax(2)", "SZmin(2),Zmax(2)",
    #     ]
    #     sweep_theta_values = [0, 80]
    #     all_data = []
    #
    #     print(f"    >> 尝试从结果树中提取 S 参数...")
    #
    #     for mode in modes_to_extract:
    #         tree_item = f"1D Results\\S-Parameters\\{mode}"
    #         try:
    #             # 获取该模式下的所有运行 ID
    #             # self.mws._FlagAsMethod("GetResultIDsFromTreeItem")
    #             run_ids = self.mws.ResultTree.GetResultIDsFromTreeItem(tree_item)
    #             print(run_ids)
    #             if run_ids is None or len(run_ids) == 0:
    #                 print(f"    [提示] 模式 {mode} 没有结果数据")
    #                 continue
    #
    #             print(f"    >> 模式 {mode}: 发现 {len(run_ids)} 个结果")
    #             res_type = None
    #             for idx, run_id in enumerate(run_ids):
    #                 if idx >= len(sweep_theta_values):
    #                     break
    #
    #                 current_theta = sweep_theta_values[idx]
    #
    #                 # 获取结果对象
    #                 self.mws._FlagAsMethod("GetResultFromTreeItem")
    #                 spare = self.mws.ResultTree.GetResultFromTreeItem(tree_item, run_id)
    #
    #
    #                 res_type = spare.GetResultObjectType()
    #
    #                 if res_type == "1DC":
    #                     try:
    #                         freq_arr = spare.GetArray("x")
    #                         s_re_arr = spare.GetArray("yre")
    #                         s_im_arr = spare.GetArray("yim")
    #
    #                         if freq_arr is None or len(freq_arr) == 0:
    #                             print(f"    [警告] theta={current_theta}° 的频率数据为空")
    #                             continue
    #
    #                         for j in range(len(freq_arr)):
    #                             c_val = complex(s_re_arr[j], s_im_arr[j])
    #                             mag_db = 20 * math.log10(abs(c_val) + 1e-12)
    #                             phase = math.degrees(math.atan2(s_im_arr[j], s_re_arr[j]))
    #
    #                             all_data.append({
    #                                 "Theta": current_theta,
    #                                 "Frequency_GHz": freq_arr[j],
    #                                 "Mode": mode,
    #                                 "Magnitude_dB": mag_db,
    #                                 "Phase_deg": phase
    #                             })
    #
    #                         print(f"    ✓ theta={current_theta}°: 提取 {len(freq_arr)} 频点")
    #                     except Exception as extract_err:
    #                         print(f"    [警告] 提取 theta={current_theta}° 数据时出错: {extract_err}")
    #                 else:
    #                     print(f"    [警告] 模式 {mode} 的结果类型为 {res_type}，期望 1DC")
    #
    #         except Exception as e:
    #             print(f"    [Warning] 提取模式 {mode} 时发生异常: {e}")
    #
    #     # 保存提取的数据
    #     if all_data:
    #         df = pd.DataFrame(all_data)
    #         df.to_csv(out_path, index=False)
    #         print(f"    ✅ 成功导出 S 参数，共 {len(df)} 行数据 -> {out_path}")
    #         return True
    #     else:
    #         print("    ❌ 未能成功提取出任何数据，请检查求解过程。")
    #         return False

    def extract_db_results(self, res_dir, out_name):
        """改进的结果提取：增强稳健性和错误处理"""
        out_path = os.path.join(res_dir, out_name)
        # 创建文件（如果不存在）
        open(out_path, "a").close()
        target_modes = {
            "S11_TE": "SZmax(1),Zmax(1)",
            "S21_TE": "SZmin(1),Zmax(1)",
            "S11_TM": "SZmax(2),Zmax(2)",
            "S21_TM": "SZmin(2),Zmax(2)"
        }

        # 【修复 1】补全所有被扫描的角度，确保 RunID 对应正确
        sweep_theta_values = [0, 20,40,60,80]
        all_data = []

        print(f"    >> 尝试从结果树中提取 S 参数...")

        # 提前获取 Resulttree 对象，避免在循环中反复调用引发异常
        try:
            rtree = self.mws.ResultTree#获取对象，返回结果为<COMObject <unknown>>
        except Exception as e:
            print("    ❌ 无法获取 CST ResultTree 对象，请确认仿真是否正常结束。")
            return False

        # 用于存储 4 个模式对应的 4 个子 DataFrame
        # 我们最后会把这 4 个表按角度和频率拼成一张大表
        all_data= []

        for mode_name, node_path in target_modes.items():
            tree_item = f"1D Results\\S-Parameters\\{node_path}"
            try:
                # 【修复 2】直接通过 rtree 调用，去除错误的 _FlagAsMethod
                run_ids = rtree.GetResultIDsFromTreeItem(tree_item)

                if run_ids is None or len(run_ids) == 0:
                    print(f"    [提示] 模式 {node_path} 没有结果数据")
                    continue

                print(f"    >> 模式 {node_path}: 发现 {len(run_ids)} 个 Sweep 结果")

                mode_data_list = []

                for idx, run_id in enumerate(run_ids[1:]):#enumerate在遍历可迭代对象时，同时得到索引（序号）和元素本身,0在cst里是当前结果，因此从1开始取
                    current_theta = sweep_theta_values[idx]

                    # 获取具体的 1D 结果对象
                    spare = rtree.GetResultFromTreeItem(tree_item, run_id)#返回<COMObject <unknown>>工程对象

                    res_type = spare.GetResultObjectType#确保获取对象为1D结果

                    if res_type == "1DC":
                        try:
                            # 读取数组数据
                            freq_arr = spare.GetArray("x")
                            s_re_arr = spare.GetArray("yre")
                            s_im_arr = spare.GetArray("yim")

                            if freq_arr is None or len(freq_arr) == 0:
                                print(f"    [警告] theta={current_theta}° 的频率数据为空")
                                continue

                            # 换算 dB 和 Phase 并存入字典
                            for j in range(len(freq_arr)):
                                c_val = complex(s_re_arr[j], s_im_arr[j])
                                mag_db = 20 * math.log10(abs(c_val) + 1e-12)
                                phase = math.degrees(math.atan2(s_im_arr[j], s_re_arr[j]))

                                mode_data_list.append({
                                    "Angle": current_theta,
                                    "Frequency_GHz": freq_arr[j],
                                    f"{mode_name}_Mag_dB": mag_db,
                                    f"{mode_name}_Phase_deg": phase
                                })
                        except Exception as extract_err:
                            print(f"    [Warning] 处理模式 {mode_name} 时发生异常: {extract_err}")
                    else:
                        print(f"    [警告] 模式 {node_path} 的结果类型为 {res_type}，期望是 1DC")
                # 将当前模式（如 S11_TE 的所有角度和频率数据）转为 DataFrame
                if mode_data_list:
                    df_temp = pdd.DataFrame(mode_data_list)
                    all_data.append(df_temp)

            except Exception as e:
                print(f"    [Warning] 处理模式 {mode_name} 时发生异常: {e}")

        # === 核心数据拼装逻辑 ===
        if not all_data:
            print("    ❌ 提取失败，所有模式均无有效数据。")
            return False

        # 以第一个模式的 DataFrame 为基准表
        final_df = all_data[0]

        # 将后续模式的表，按照 Angle 和 Frequency 左右合并对齐 (Merge)
        # 这样就能实现 TE 和 TM 数据在同一行的完美并排
        for i in range(1, len(all_data)):
            final_df = pdd.merge(final_df, all_data[i], on=['Angle', 'Frequency_GHz'], how='outer')

        # 最后，按照先 Angle 后 Frequency 的顺序排序，确保 0度、20度... 按顺序往下排
        final_df.sort_values(by=['Angle', 'Frequency_GHz'], inplace=True)

        # 整理表头顺序，看起来更舒服
        cols_order = ['Angle', 'Frequency_GHz',
                      'S11_TE_Mag_dB', 'S11_TE_Phase_deg', 'S21_TE_Mag_dB', 'S21_TE_Phase_deg',
                      'S11_TM_Mag_dB', 'S11_TM_Phase_deg', 'S21_TM_Mag_dB', 'S21_TM_Phase_deg']

        # 防止某些模式没提取到导致报错，只选取存在的列
        valid_cols = [c for c in cols_order if c in final_df.columns]
        final_df = final_df[valid_cols]

        # 导出为 CSV
        final_df.to_csv(out_path, index=False)
        print(f"    ✅ 成功聚合全极化 S 参数，导出 {len(final_df)} 行完美矩阵 -> {out_path}")
        return True
class Material(StructureMacros):
    MaterialName = ""
    MaterialEpsilon = 0
    MaterialMu = 0
    Classname = "Material"

    def __init__(self, handle, Name, Epsilon, Mu) -> None:
        super().__init__(handle)
        self.MaterialName = Name
        self.MaterialEpsilon = Epsilon
        self.MaterialMu = Mu
        self.materialcreate()

    def materialcreate(self):
        sCommand = f"""With Material 
     .Reset 
     .Name "{self.MaterialName}"
     .Folder ""
     .Rho "0.0"
     .ThermalType "Normal"
     .ThermalConductivity "0"
     .SpecificHeat "0", "J/K/kg"
     .DynamicViscosity "0"
     .Emissivity "0"
     .MetabolicRate "0.0"
     .VoxelConvection "0.0"
     .BloodFlow "0"
     .MechanicsType "Unused"
     .IntrinsicCarrierDensity "0"
     .FrqType "all"
     .Type "Normal"
     .MaterialUnit "Frequency", "GHz"
     .MaterialUnit "Geometry", "mm"
     .MaterialUnit "Time", "ns"
     .MaterialUnit "Temperature", "Kelvin"
     .Epsilon "{self.MaterialEpsilon}"
     .Mu "{self.MaterialMu}"
     .Sigma "0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .SetConstTanDStrategyEps "AutomaticOrder"
     .ConstTanDModelOrderEps "3"
     .DjordjevicSarkarUpperFreqEps "0"
     .SetElParametricConductivity "False"
     .ReferenceCoordSystem "Global"
     .CoordSystemType "Cartesian"
     .SigmaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .SetConstTanDStrategyMu "AutomaticOrder"
     .ConstTanDModelOrderMu "3"
     .DjordjevicSarkarUpperFreqMu "0"
     .SetMagParametricConductivity "False"
     .DispModelEps  "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .MaximalOrderNthModelFitEps "10"
     .ErrorLimitNthModelFitEps "0.1"
     .UseOnlyDataInSimFreqRangeNthModelEps "False"
     .DispersiveFittingSchemeMu "Nth Order"
     .MaximalOrderNthModelFitMu "10"
     .ErrorLimitNthModelFitMu "0.1"
     .UseOnlyDataInSimFreqRangeNthModelMu "False"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .NLAnisotropy "False"
     .NLAStackingFactor "1"
     .NLADirectionX "1"
     .NLADirectionY "0"
     .NLADirectionZ "0"
     .Colour "0", "1", "0" 
     .Wireframe "False" 
     .Reflection "False" 
     .Allowoutline "True" 
     .Transparentoutline "False" 
     .Transparency "0" 
     .Create
    End With"""
        self.AddToHistoryWithCommand(
            Tag="Add Material " + self.MaterialName, Command=sCommand
        )
        return self


class GeneralModel(StructureMacros):
    Component = ""
    Name = ""
    Material = ""
    Tag = ""
    Classname = "GeneralModel"

    def __init__(self, handle, Tag, Component, Name, Material) -> None:
        super().__init__(handle)
        self.Component = Component
        self.Name = Name
        self.Material = Material
        self.Tag = Tag

    def create(self):
        pass


class Brick(GeneralModel):
    Xrange = [0, 1]
    Yrange = [0, 1]
    Zrange = [0, 1]
    Component = "Hallo"
    Name = "World"
    Material = "PEC"
    Classname = "Brick"

    def __init__(
        self, handle, Tag, Component, Name, Material, Xrange, Yrange, Zrange
    ) -> None:
        super().__init__(handle, Tag, Component, Name, Material)
        self.Xrange = Xrange
        self.Yrange = Yrange
        self.Zrange = Zrange
        self.create()

    def create(self):
        Command = f"""With Brick
     .Reset 
     .Name "{self.Name}" 
     .Component "{self.Component}" 
     .Material "{self.Material}" 
     .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}" 
     .Yrange "{self.Yrange[0]}", "{self.Yrange[1]}"
     .Zrange "{self.Zrange[0]}", "{self.Zrange[1]}" 
     .Create
    End With"""
        self.AddToHistoryWithCommand(self.Tag, Command)
        return self


class Cylinder(GeneralModel):
    Material = "Vacuum"
    Innerradius = 0
    Outerradius = 0
    Xcenter = 0
    Ycenter = 0
    Zcenter = 0
    Xrange = [0, 0]
    Yrange = [0, 0]
    Zrange = [0, 0]
    Range = [0, 0]
    Segments = 0
    Axis = "z"
    Classname = "Cylinder"

    def __init__(
        self,
        handle,
        Tag,
        Component,
        Name,
        Material,
        Axis,
        Innerradius,
        Outerradius,
        Center,
        Range,
        Segments=0,
    ) -> None:
        super().__init__(handle, Tag, Component, Name, Material)
        self.Innerradius = Innerradius
        self.Outerradius = Outerradius
        self.Xcenter = Center[0]
        self.Ycenter = Center[1]
        self.Zcenter = Center[2]
        self.Range = Range
        self.Segments = Segments
        self.Axis = Axis
        self.create()

    def create(self):
        sCommand = f"""With Cylinder
    .Reset
    .Name ("{self.Name}")
    .Component ("{self.Component}")
    .Material ("{self.Material}")
    .Axis ("{self.Axis}")
    .Outerradius ("{self.Outerradius}")
    .Innerradius ("{self.Innerradius}")
    .Xcenter ("{self.Xcenter}")
    .Ycenter ("{self.Ycenter}")
    .Zcenter ("{self.Zcenter}")"""

        if self.Axis == "z":
            sCommand = (
                sCommand
                + f"""
    .Zrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
    End With"""
            )
        elif self.Axis == "y":
            sCommand = (
                sCommand
                + f"""
    .Yrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
    End With"""
            )
        elif self.Axis == "x":
            sCommand = (
                sCommand
                + f"""
    .Xrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
    End With"""
            )
        self.AddToHistoryWithCommand(Tag=self.Tag, Command=sCommand)
        return self


class AnalyticalFace(GeneralModel):
    LawX = ""
    LawY = ""
    LawZ = ""
    ParameterRangeU = ""
    ParameterRangeV = ""
    Classname = "AnalyticalFace"

    def __init__(
        self,
        handle,
        Tag,
        Component,
        Name,
        Material,
        LawX,
        LawY,
        LawZ,
        ParameterRangeU,
        ParameterRangeV,
    ) -> None:
        super().__init__(handle, Tag, Component, Name, Material)
        self.LawX = LawX
        self.LawY = LawY
        self.LawZ = LawZ
        self.ParameterRangeU = ParameterRangeU
        self.ParameterRangeV = ParameterRangeV
        self.create()

    def create(self):
        sCommand = f"""
        With AnalyticalFace
            .Reset 
            .Name "{self.Name}" 
            .Component "{self.Component}" 
            .Material "{self.Material}" 
            .LawX "{self.LawX}" 
            .LawY "{self.LawY}" 
            .LawZ "{self.LawZ}" 
            .ParameterRangeU "{self.ParameterRangeU[0]}", "{self.ParameterRangeU[1]}" 
            .ParameterRangeV "{self.ParameterRangeV[0]}", "{self.ParameterRangeV[1]}" 
            .Create
        End With"""
        self.AddToHistoryWithCommand(self.Tag, sCommand)
        return super().create()


class Pick(StructureMacros):
    Classname = "Pick"

    def __init__(self, handle) -> None:
        super().__init__(handle)

    def PickCenterpointFromId(self, Tag, Component, Name, Id):
        sCommand = f'Pick.PickCenterpointFromId "{Component}:{Name}", "{Id}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickFaceFromId(self, Tag, Component, Name, Id):
        sCommand = f'Pick.PickFaceFromId "{Component}:{Name}", "{Id}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickEdgeFromId(self, Tag, Component, Name, edge_id, vertex_id):
        sCommand = (
            f'Pick.PickEdgeFromId "{Component}:{Name}", "{edge_id}", "{vertex_id}"'
        )
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickEdgeFromPoint(self, Tag, Component, Name, PointCoordinate):
        sCommand = f'Pick.PickEdgeFromPoint "{Component}:{Name}",{PointCoordinate[0]},{PointCoordinate[1]},{PointCoordinate[2]}'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickFaceFromPoint(self, Tag, Component, Name, PointCoordinate):
        sCommand = f'Pick.PickFaceFromPoint "{Component}:{Name}",{PointCoordinate[0]},{PointCoordinate[1]},{PointCoordinate[2]}'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickPointFromCoordinates(self, Tag, PointCoordinate):
        sCommand = f"Pick.PickPointFromCoordinates {PointCoordinate[0]},{PointCoordinate[1]},{PointCoordinate[2]}"
        self.AddToHistoryWithCommand(Tag, sCommand)


class WCS(StructureMacros):
    Classname = "WCS"

    def __init__(self, handle) -> None:
        super().__init__(handle)

    def AlignWCSWithSelectedPoint(self, Tag):
        self.AddToHistoryWithCommand(Tag, 'WCS.AlignWCSWithSelected "Point"')

    def ActivateWCSGlobal(self):
        self.AddToHistoryWithCommand(
            "Active the global WCS", 'WCS.ActivateWCS "global"'
        )

    def AlignWCSWithSelectedFace(self, Tag):
        self.AddToHistoryWithCommand(Tag, 'WCS.AlignWCSWithSelected "Face"')


class Transform(StructureMacros):
    Classname = "Transform"

    def __init__(self, handle) -> None:
        super().__init__(handle)

    def MirrorTransForm(self, Tag, Component, Name, NormalVector, Copy):
        sCommand = f"""With Transform 
     .Reset 
     .Name "{Component}:{Name}" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "{NormalVector[0]}", "{NormalVector[1]}", "{NormalVector[2]}" 
     .MultipleObjects "{Copy}" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .Transform "Shape", "Mirror" 
    End With"""
        self.AddToHistoryWithCommand(Tag, sCommand)

    def TranslateTransform(
        self, Tag, Component, Name, TransVector, Copy, RepetitionFactor
    ):
        sCommand = f"""With Transform 
        .Reset 
        .Name "{Component}:{Name}" 
        .Vector "{TransVector[0]}", "{TransVector[1]}", "{TransVector[2]}" 
        .UsePickedPoints "False" 
        .InvertPickedPoints "False" 
        .MultipleObjects "{Copy}" 
        .GroupObjects "False" 
        .Repetitions "{RepetitionFactor}" 
        .MultipleSelection "False" 
        .Destination "" 
        .Material "" 
        .Transform "Shape", "Translate" 
    End With"""
        self.AddToHistoryWithCommand(Tag, sCommand)


class Solid(StructureMacros):
    Classname = "Solid"

    def __init__(self, handle) -> None:
        super().__init__(handle)

    def Subtract(self, Tag, component1, name1, component2, name2):
        sCommand = f'Solid.Subtract "{component1}:{name1}", "{component2}:{name2}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def Add(self, Tag, component1, name1, component2, name2):
        sCommand = f'Solid.Add "{component1}:{name1}", "{component2}:{name2}"'
        self.AddToHistoryWithCommand(Tag, sCommand)  # 留下来的是后面的那个面，后来居上

    def BlendEdge(self, Tag, radius):
        sCommand = f'Solid.BlendEdge "{radius}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def Insert(self, Tag, component1, name1, component2, name2):
        sCommand = f'Solid.Insert "{component1}:{name1}", "{component2}:{name2}"'
        self.AddToHistoryWithCommand(Tag, sCommand)


class Loft(GeneralModel):
    Classname = "Loft"
    Tangency = "0"

    def __init__(self, handle, Tag, Component, Name, Material, Tangency) -> None:
        super().__init__(handle, Tag, Component, Name, Material)
        self.Tangency = Tangency
        self.create()

    def create(self):
        sCommand = f""" 
        With Loft
            .Reset
            .Name "{self.Name}"
            .Component "{self.Component}"
            .Material "{self.Material}"
            .Tangency "{self.Tangency}"
            .CreateNew
        End With"""
        self.AddToHistoryWithCommand(self.Tag, sCommand)
        return super().create()


class Extrude(GeneralModel):
    Mode = "Picks"
    Height = ""
    Twist = ""
    Taper = ""
    UsePicksForHeight = "False"
    DeleteBaseFaceSolid = "False"
    KeepMaterial = "False"
    ClearPickedFace = "True"

    def __init__(
        self, handle, Tag, Component, Name, Material, Height, Twist, Taper, **kwargs
    ) -> None:
        super().__init__(handle, Tag, Component, Name, Material)
        self.Height = Height
        self.Twist = Twist
        self.Taper = Taper
        for key, value in kwargs.items():
            match key:
                case "Mode":
                    self.Mode = value
                case "UsePicksForHeight":
                    self.UsePicksForHeight = value
                case "DeleteBaseFaceSolid":
                    self.DeleteBaseFaceSolid = value
                case "KeepMaterial":
                    self.KeepMaterial = value
                case "ClearPickedFace":
                    self.ClearPickedFace = value
                case _:
                    raise ("DameDane")
        self.Tag = Tag
        self.create()

    def create(self):
        sCommand = f"""
        With Extrude 
            .Reset 
            .Name "{self.Name}" 
            .Component "{self.Component}" 
            .Material "{self.Material}" 
            .Mode "{self.Mode}" 
            .Height "{self.Height}" 
            .Twist "{self.Twist}" 
            .Taper "{self.Taper}" 
            .UsePicksForHeight "{self.UsePicksForHeight}" 
            .DeleteBaseFaceSolid "{self.DeleteBaseFaceSolid}" 
            .KeepMaterials "{self.KeepMaterial}" 
            .ClearPickedFace "{self.ClearPickedFace}" 
            .Create 
        End With
        """
        self.AddToHistoryWithCommand(self.Tag, sCommand)
        return super().create()


class Port(StructureMacros):
    PortNumber = 1
    NumberOfModes = 1
    Coordinates = "Picks"
    Orientation = "positive"
    PortOnBound = "True"
    AdjustPolarization = "False"
    Xrange = [0, 0]
    XrangeAdd = [0, 0]
    Yrange = [0, 0]
    YrangeAdd = [0, 0]
    Zrange = [0, 0]
    ZrangeAdd = [0, 0]
    Classname = "Port"

    def __init__(self, handle, Tag, Range, PortNumber, **kwargs) -> None:
        super().__init__(handle)
        self.Xrange = Range[0]
        self.Yrange = Range[1]
        self.Zrange = Range[2]
        self.PortNumber = PortNumber
        for key, value in kwargs.items():
            match key:
                case "NumberOfModes":
                    self.NumberOfModes = value
                case "Coordinates":
                    self.Coordinates = value
                case "Orientation":
                    self.Orientation = value
                case "PortOnBound":
                    self.PortOnBound = value
                case "AdjustPolarization":
                    self.AdjustPolarization = value
                case "AddRange":
                    self.XrangeAdd = [0]
                    self.YrangeAdd = [1]
                    self.ZrangeAdd = [2]
        self.create(Tag)

    def create(self, Tag):
        self.Tag = Tag
        sCommand = f"""
    With Port 
        .Reset 
        .PortNumber "{self.PortNumber}" 
        .Label ""
        .Folder ""
        .NumberOfModes "{self.NumberOfModes}"
        .AdjustPolarization "{self.AdjustPolarization}"
        .PolarizationAngle "0.0"
        .ReferencePlaneDistance "0"
        .TextSize "50"
        .TextMaxLimit "0"
        .Coordinates "{self.Coordinates}"
        .Orientation "{self.Orientation}"
        .PortOnBound "{self.PortOnBound}"
        .ClipPickedPortToBound "False"
        .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}"
        .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}"
        .Yrange "{self.Zrange[0]}", "{self.Zrange[1]}"
        .XrangeAdd "{self.XrangeAdd[0]}", "{self.XrangeAdd[1]}"
        .XrangeAdd "{self.XrangeAdd[0]}", "{self.XrangeAdd[1]}"
        .ZrangeAdd "{self.ZrangeAdd[0]}", "{self.ZrangeAdd[1]}"
        .SingleEnded "False"
        .WaveguideMonitor "False"
        .Create 
    End With"""
        self.AddToHistoryWithCommand(
            self.Tag + "Add Port" + str(self.PortNumber), sCommand
        )
        return self


class Mesh(StructureMacros):
    StepsPerWaveNear = 17
    StepsPerWaveFar = 10
    StepsPerBoxNear = 12
    StepsPerBoxFar = 10
    MeshType = "Tetrahedral"
    SetCreator = "High Frequency"

    def __init__(self, handle) -> None:
        self.mws = handle

    def init(
        self,
        StepsPerWaveNear,
        StepsPerWaveFar,
        StepsPerBoxNear,
        StepsPerBoxFar,
        **kwargs,
    ):
        self.StepsPerBoxNear = StepsPerBoxNear
        self.StepsPerBoxFar = StepsPerBoxFar
        self.StepsPerWaveNear = StepsPerWaveNear
        self.StepsPerWaveFar = StepsPerWaveFar
        for key, value in kwargs.items():
            match key:
                case "MeshType":
                    self.MeshType = value
                case "SetCreator":
                    self.SetCreator = value
        return self

    def MeshUpdate(self, Tag):
        sCommand = f"""
        With Mesh 
            .MeshType "{self.MeshType}" 
            .SetCreator "{self.SetCreator}"
        End With 
        With MeshSettings 
            'MAX CELL - WAVELENGTH REFINEMENT 
            .Set "StepsPerWaveNear", "{self.StepsPerWaveNear}" 
            .Set "StepsPerWaveFar", "{self.StepsPerWaveFar}" 
            .Set "PhaseErrorNear", "0.02" 
            .Set "PhaseErrorFar", "0.02" 
            .Set "CellsPerWavelengthPolicy", "cellsperwavelength" 
            'MAX CELL - GEOMETRY REFINEMENT 
            .Set "StepsPerBoxNear", "{self.StepsPerBoxNear}" 
            .Set "StepsPerBoxFar", "{self.StepsPerBoxFar}" 
            .Set "ModelBoxDescrNear", "maxedge" 
            .Set "ModelBoxDescrFar", "maxedge" 
            'MIN CELL 
            .Set "UseRatioLimit", "0" 
            .Set "RatioLimit", "100" 
            .Set "MinStep", "0" 
            'MESHING METHOD 
            .SetMeshType "Unstr" 
            .Set "Method", "0" 
        End With 
        With MeshSettings 
            .SetMeshType "Tet" 
            .Set "CurvatureOrder", "1" 
            .Set "CurvatureOrderPolicy", "automatic" 
            .Set "CurvRefinementControl", "NormalTolerance" 
            .Set "NormalTolerance", "22.5" 
            .Set "SrfMeshGradation", "1.5" 
            .Set "SrfMeshOptimization", "1" 
        End With 
        With MeshSettings 
            .SetMeshType "Unstr" 
            .Set "UseMaterials",  "1" 
            .Set "MoveMesh", "0" 
        End With 
        With MeshSettings 
            .SetMeshType "All" 
            .Set "AutomaticEdgeRefinement",  "0" 
        End With 
        With MeshSettings 
            .SetMeshType "Tet" 
            .Set "UseAnisoCurveRefinement", "1" 
            .Set "UseSameSrfAndVolMeshGradation", "1" 
            .Set "VolMeshGradation", "1.5" 
            .Set "VolMeshOptimization", "1" 
        End With 
        With MeshSettings 
            .SetMeshType "Unstr" 
            .Set "SmallFeatureSize", "0" 
            .Set "CoincidenceTolerance", "1e-06" 
            .Set "SelfIntersectionCheck", "1" 
            .Set "OptimizeForPlanarStructures", "0" 
        End With 
        With Mesh 
            .SetParallelMesherMode "Tet", "maximum" 
            .SetMaxParallelMesherThreads "Tet", "1" 
        End With
        """
        self.AddToHistoryWithCommand(Tag, sCommand)

    def MeshUpdateHex(self, Tag):
        sCommand = f"""With Mesh 
                .MeshType "PBA" 
                .SetCreator "High Frequency"
            End With 
            With MeshSettings 
                .SetMeshType "Hex" 
                .Set "Version", 1%
                'MAX CELL - WAVELENGTH REFINEMENT 
                .Set "StepsPerWaveNear", "{self.StepsPerWaveNear}" 
                .Set "StepsPerWaveFar", "{self.StepsPerWaveFar}" 
                .Set "WavelengthRefinementSameAsNear", "1" 
                'MAX CELL - GEOMETRY REFINEMENT 
                .Set "StepsPerBoxNear", "{self.StepsPerBoxNear}" 
                .Set "StepsPerBoxFar", "{self.StepsPerBoxFar}" 
                .Set "MaxStepNear", "0" 
                .Set "MaxStepFar", "0" 
                .Set "ModelBoxDescrNear", "maxedge" 
                .Set "ModelBoxDescrFar", "maxedge" 
                .Set "UseMaxStepAbsolute", "0" 
                .Set "GeometryRefinementSameAsNear", "0" 
                'MIN CELL 
                .Set "UseRatioLimitGeometry", "1" 
                .Set "RatioLimitGeometry", "15" 
                .Set "MinStepGeometryX", "0" 
                .Set "MinStepGeometryY", "0" 
                .Set "MinStepGeometryZ", "0" 
                .Set "UseSameMinStepGeometryXYZ", "1" 
            End With 
            With MeshSettings 
                .Set "PlaneMergeVersion", "2" 
            End With 
            With MeshSettings 
                .SetMeshType "Hex" 
                .Set "FaceRefinementOn", "0" 
                .Set "FaceRefinementPolicy", "2" 
                .Set "FaceRefinementRatio", "2" 
                .Set "FaceRefinementStep", "0" 
                .Set "FaceRefinementNSteps", "2" 
                .Set "EllipseRefinementOn", "0" 
                .Set "EllipseRefinementPolicy", "2" 
                .Set "EllipseRefinementRatio", "2" 
                .Set "EllipseRefinementStep", "0" 
                .Set "EllipseRefinementNSteps", "2" 
                .Set "FaceRefinementBufferLines", "3" 
                .Set "EdgeRefinementOn", "1" 
                .Set "EdgeRefinementPolicy", "1" 
                .Set "EdgeRefinementRatio", "2" 
                .Set "EdgeRefinementStep", "0" 
                .Set "EdgeRefinementBufferLines", "3" 
                .Set "RefineEdgeMaterialGlobal", "0" 
                .Set "RefineAxialEdgeGlobal", "0" 
                .Set "BufferLinesNear", "3" 
                .Set "UseDielectrics", "1" 
                .Set "EquilibrateOn", "0" 
                .Set "Equilibrate", "1.5" 
                .Set "IgnoreThinPanelMaterial", "0" 
            End With 
            With MeshSettings 
                .SetMeshType "Hex" 
                .Set "SnapToAxialEdges", "1"
                .Set "SnapToPlanes", "1"
                .Set "SnapToSpheres", "1"
                .Set "SnapToEllipses", "1"
                .Set "SnapToCylinders", "1"
                .Set "SnapToCylinderCenters", "1"
                .Set "SnapToEllipseCenters", "1"
            End With 
            With Mesh 
                .ConnectivityCheck "True"
                .UsePecEdgeModel "True" 
                .PointAccEnhancement "0" 
                .TSTVersion "0"
                .PBAVersion "2022060322" 
                .SetCADProcessingMethod "MultiThread22", "-1" 
                .SetGPUForMatrixCalculationDisabled "False" 
            End With
            """
        self.AddToHistoryWithCommand(Tag, sCommand)
        pass


class Solver(StructureMacros):
    def __init__(self, handle) -> None:
        self.mws = handle

    def FDSolver(self):
        sCommand = """
        Mesh.SetCreator "High Frequency" 

        With FDSolver
            .Reset 
            .SetMethod "Tetrahedral", "General purpose" 
            .OrderTet "Second" 
            .OrderSrf "First" 
            .Stimulation "All", "All" 
            .ResetExcitationList 
            .AutoNormImpedance "False" 
            .NormingImpedance "50" 
            .ModesOnly "False" 
            .ConsiderPortLossesTet "True" 
            .SetShieldAllPorts "False" 
            .AccuracyHex "1e-6" 
            .AccuracyTet "1e-5" 
            .AccuracySrf "1e-3" 
            .LimitIterations "False" 
            .MaxIterations "0" 
            .SetCalcBlockExcitationsInParallel "True", "True", "" 
            .StoreAllResults "False" 
            .StoreResultsInCache "False" 
            .UseHelmholtzEquation "True" 
            .LowFrequencyStabilization "True" 
            .Type "Direct" 
            .MeshAdaptionHex "False" 
            .MeshAdaptionTet "True" 
            .AcceleratedRestart "True" 
            .FreqDistAdaptMode "Distributed" 
            .NewIterativeSolver "True" 
            .TDCompatibleMaterials "False" 
            .ExtrudeOpenBC "False" 
            .SetOpenBCTypeHex "Default" 
            .SetOpenBCTypeTet "Default" 
            .AddMonitorSamples "True" 
            .CalcPowerLoss "True" 
            .CalcPowerLossPerComponent "False" 
            .StoreSolutionCoefficients "True" 
            .UseDoublePrecision "False" 
            .UseDoublePrecision_ML "True" 
            .MixedOrderSrf "False" 
            .MixedOrderTet "False" 
            .PreconditionerAccuracyIntEq "0.15" 
            .MLFMMAccuracy "Default" 
            .MinMLFMMBoxSize "0.3" 
            .UseCFIEForCPECIntEq "True" 
            .UseEnhancedCFIE2 "True" 
            .UseFastRCSSweepIntEq "false" 
            .UseSensitivityAnalysis "False" 
            .UseEnhancedNFSImprint "False" 
            .RemoveAllStopCriteria "Hex"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
            .RemoveAllStopCriteria "Tet"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
            .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
            .RemoveAllStopCriteria "Srf"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
            .SweepMinimumSamples "3" 
            .SetNumberOfResultDataSamples "5001" 
            .SetResultDataSamplingMode "Automatic" 
            .SweepWeightEvanescent "1.0" 
            .AccuracyROM "1e-4" 
            .AddSampleInterval "", "", "1", "Automatic", "True" 
            .AddSampleInterval "", "", "", "Automatic", "False" 
            .MPIParallelization "False"
            .UseDistributedComputing "False"
            .NetworkComputingStrategy "RunRemote"
            .NetworkComputingJobCount "3"
            .UseParallelization "True"
            .MaxCPUs "1024"
            .MaximumNumberOfCPUDevices "2"
        End With

        With IESolver
            .Reset 
            .UseFastFrequencySweep "True" 
            .UseIEGroundPlane "False" 
            .SetRealGroundMaterialName "" 
            .CalcFarFieldInRealGround "False" 
            .RealGroundModelType "Auto" 
            .PreconditionerType "Auto" 
            .ExtendThinWireModelByWireNubs "False" 
            .ExtraPreconditioning "False" 
        End With

        With IESolver
            .SetFMMFFCalcStopLevel "0" 
            .SetFMMFFCalcNumInterpPoints "6" 
            .UseFMMFarfieldCalc "True" 
            .SetCFIEAlpha "0.500000" 
            .LowFrequencyStabilization "False" 
            .LowFrequencyStabilizationML "True" 
            .Multilayer "False" 
            .SetiMoMACC_I "0.0001" 
            .SetiMoMACC_M "0.0001" 
            .DeembedExternalPorts "True" 
            .SetOpenBC_XY "True" 
            .OldRCSSweepDefintion "False" 
            .SetRCSOptimizationProperties "True", "100", "0.00001" 
            .SetAccuracySetting "Custom" 
            .CalculateSParaforFieldsources "True" 
            .ModeTrackingCMA "True" 
            .NumberOfModesCMA "3" 
            .StartFrequencyCMA "-1.0" 
            .SetAccuracySettingCMA "Default" 
            .FrequencySamplesCMA "0" 
            .SetMemSettingCMA "Auto" 
            .CalculateModalWeightingCoefficientsCMA "True" 
            .DetectThinDielectrics "True" 
        End With
        """
        self.AddToHistoryWithCommand("设置求解器", sCommand)


# def CstSaveAsProject(mws, projectName):
#     mws._FlagAsMethod("SaveAs")
#     mws.SaveAs(projectName, "false")


class PostProcessingItems(StructureMacros):
    def __init__(self, handle) -> None:
        self.mws = handle
        pass

    def GetSparametersinRunID(self, ResultTag="S11"):
        match ResultTag:
            case "S11":
                TreeItem = "1D Results\\S-Parameters\\S1,1"  # (python记得写双斜杠哦)
            case "S12":
                TreeItem = "1D Results\\S-Parameters\\S1,2"
            case "S21":
                TreeItem = "1D Results\\S-Parameters\\S1,2"
            case "S22":
                TreeItem = "1D Results\\S-Parameters\\S2,2"

        resultdatas = []
        SREALseries = []
        SIMAGEseries = []
        Frequencyseries = []

        # 'get an array of existing result ids for this tree item
        IDs = self.mws.ResultTree.GetResultIDsFromTreeItem(
            TreeItem
        )  # 返回的是Result Navigator里面的RunID，如果有扫参的话就会有不同的ID出现
        for N in range(len(IDs)):
            spara = self.mws.ResultTree.GetResultFromTreeItem(TreeItem, IDs[N])
            # GetResultObjectType可不能在后面加上括号，因为在Result1DComplex Object里面的这个方法就没有括号
            resulttype = spara.GetResultObjectType
            if resulttype == "1DC":
                resultdatas.append(spara)
                FrequencyRange = spara.GetArray("x")
                Frequencyseries.append(FrequencyRange)
                SRE = spara.GetArray("yre")
                SREALseries.append(SRE)
                SIM = spara.GetArray("yim")
                SIMAGEseries.append(SIM)
                # plt.plot(FrequencyRange, SRE, label='RealPart')
                # plt.plot(FrequencyRange, SIM, label='Imag Part')
                # plt.xlabel(spara.GetXlabel)
                # plt.ylabel(spara.GetYlabel+ResultTag)
                # plt.title('Current Schematic:RunID is '+IDs[N])
                # plt.show()  # 这些都是测试板块啦
        return resultdatas, Frequencyseries, SREALseries, SIMAGEseries

    def GetSelectedTreeItem(self):
        # 返回一个已经打开的项目的选中的工程树中的命令的路径地址
        # 需要注意的是python的路径并不能用单斜杠表示，需要在路径的前方加上r进行转义或者使用双斜杠
        select_item_path = self.mws.GetSelectedTreeItem
        select_item_path = select_item_path.replace("\\", "\\\\")
        return select_item_path


class Monitor(StructureMacros):
    def __init__(self, handle) -> None:
        self.mws = handle
        pass

    def CreateUsingArbitraryValues(self, Tag, frequencylist):
        freqstr = ""
        for freq in frequencylist:
            freqstr = freqstr + str(freq) + ";"
            pass

        sCommand = f"""
        With Monitor
                .Reset
                .Domain "Frequency"
                .FieldType "Efield"
                .Dimension "Volume"
                .Coordinates "Structure"
                .CreateUsingArbitraryValues "{freqstr}"
        End With
        """
        self.AddToHistoryWithCommand(Tag, sCommand)

    def CreateUsingLinearSamples(self, Tag, FreqencyRange, Samples):
        sCommand = f"""With Monitor
          .Reset 
          .Domain "Frequency"
          .FieldType "Efield"
          .Dimension "Volume" 
          .Coordinates "Structure" 
          .CreateUsingLinearSamples "{FreqencyRange[0]}", "{FreqencyRange[1]}", "{Samples}"
        End With
        """
        self.AddToHistoryWithCommand(Tag, sCommand)
        pass


class FDSolver(StructureMacros):
    def __init__(self, handle) -> None:
        self.mws = handle

    def FDSolverSetting(self, meshtype, broadbandsweepmethod):
        if meshtype == "Hexahedral" and broadbandsweepmethod == "General purpose":
            sCommand = """Mesh.SetCreator "High Frequency" 
                With FDSolver
                    .Reset 
                    .SetMethod "Hexahedral", "General purpose" 
                End With
                """
            self.AddToHistoryWithCommand("Set Mesh with Hexahedral", sCommand)
        else:
            raise ("NOSUCHTEMPLATE")
