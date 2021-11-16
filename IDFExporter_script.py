__doc__ = "Energy Analisys"
__author__ = 'Lucas Vieira'
__title__ = "IDF Exporter"

#importing revit libraries
from os import remove
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Analysis import *
from Autodesk.Revit.DB.UnitTypeId import *

#importing system libraries
import clr
import System
import sys
import ctypes
import os.path
import operator

import xml.dom.minidom

#atributting the active window in Revit as the model source
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#creating object classes for Energy Analysis on Energy Plus

class Version:
    def __init__(self, ClassName, VersionIdentifier):
        self.ClassName = ClassName
        self.VersionIdentifier = VersionIdentifier

class SimulationControl:
    def __init__(self, ClassName, DoZoneSizingCalculation, DoSystemSizingCalculation, DoPlantSizingCalculation, RunSimulationForSizingPeriods,\
    RunSimulationForWeatherFileRunPeriods, DoHVACSizingSimualation, MaxNumberOfHVACSizingSimulation):
        self.ClassName = ClassName
        self.DoZoneSizingCalculation = DoZoneSizingCalculation
        self.DoSystemSizingCalculation = DoSystemSizingCalculation
        self.DoPlantSizingCalculation = DoPlantSizingCalculation
        self.RunSimulationForSizingPeriods = RunSimulationForSizingPeriods
        self.RunSimulationForWeatherFileRunPeriods = RunSimulationForWeatherFileRunPeriods
        self.DoHVACSizingSimualation = DoHVACSizingSimualation
        self.MaxNumberOfHVACSizingSimulation = MaxNumberOfHVACSizingSimulation
     
class Building:
    def __init__(self,ClassName, Name, NorthAxis, Terrain, W, DeltaC, SolarDistribution, MaxNumberOfWarmUpDays, MinNumberOfWarmUpDays):
        self.ClassName = ClassName
        self.Name = Name
        self.NorthAxis = NorthAxis
        self.Terrain = Terrain
        self.W = W
        self.DeltaC = DeltaC
        self.SolarDistribution = SolarDistribution
        self.MaxNumberOfWarmUpDays = MaxNumberOfWarmUpDays
        self.MinNumberOfWarmUpDays = MinNumberOfWarmUpDays

class SurfaceConvectionAlgorithm:
    def __init__(self,ClassName,Algorithm):
        self.ClassName = ClassName
        self.Algorithm = Algorithm

class HeatBalanceAlgorithm:
    def __init__(self,ClassName,Algorithm, SurfaceTemperatureUpperLimit, MinSurfaceConvectionHeatTransferCoefficientValue, MaxSurfaceConvectionHeatTransferCoefficientValue):
        self.ClassName = ClassName
        self.Algorithm = Algorithm
        self.SurfaceTemperatureUpperLimit = SurfaceTemperatureUpperLimit
        self.MinSurfaceConvectionHeatTransferCoefficientValue = MinSurfaceConvectionHeatTransferCoefficientValue
        self.MaxSurfaceConvectionHeatTransferCoefficientValue = MaxSurfaceConvectionHeatTransferCoefficientValue

class Timestep:
    def __init__(self,ClassName, NumberOfTimestepsPerHour):
        self.ClassName = ClassName
        self.NumberOfTimestepsPerHour = NumberOfTimestepsPerHour

class Site:
    def __init__(self,ClassName, Name, Latitude, Longitude, TimeZone, Elevation):
        self.ClassName = ClassName
        self.Name = Name
        self.Latitude = Latitude
        self.Longitude = Longitude
        self.TimeZone = TimeZone
        self.Elevation = Elevation

class SizingPeriod:
    def __init__(self,ClassName, Name, Month, DayOfMonth, DayType, MaximumDryBulbTemperature, DailyDryBulbTemperatureRange, DryBulbTemperatureRangeModifierType, DryBulbTemperatureRangeModifierDayScheduleName, HumidityConditionType, WetBulbOrDewPointAtMaximumDryBulb, HumidityConditionDayScheduleName, HumidityRatioAtMaximumDryBulb, \
        EnthalpyAtMaximumDryBulb, DailyWetBulbTemperatureRange, BarometricPressure, WindSpeed, WindDirection, RainIndicator, SnowIndicator, DaylightSavingTimeIndicator, SolarModelIndicator, BeamSolarDayScheduleName, DiffuseSolarDayScheduleName, \
        ASHRAEClearSkyOpticalForBeamIrradiance, ASHRAEClearSkyOpticalForBDiffuseIrradiance, SkyClearness, MaximumNumberWarmUpDays,BeginEnvironmentResetMode):
        self.ClassName = ClassName
        self.Name = Name
        self.Month = Month
        self.DayOfMonth = DayOfMonth
        self.DayType = DayType
        self.MaximumDryBulbTemperature = MaximumDryBulbTemperature
        self.DailyDryBulbTemperatureRange = DailyDryBulbTemperatureRange
        self.DryBulbTemperatureRangeModifierType = DryBulbTemperatureRangeModifierType
        self.DryBulbTemperatureRangeModifierDayScheduleName = DryBulbTemperatureRangeModifierDayScheduleName
        self.HumidityConditionType = HumidityConditionType
        self.HWetBulbOrDewPointAtMaximumDryBulb = WetBulbOrDewPointAtMaximumDryBulb
        self.HumidityConditionDayScheduleName = HumidityConditionDayScheduleName
        self.HumidityRatioAtMaximumDryBulb = HumidityRatioAtMaximumDryBulb
        self.EnthalpyAtMaximumDryBulb = EnthalpyAtMaximumDryBulb
        self.DailyWetBulbTemperatureRange = DailyWetBulbTemperatureRange
        self.HBarometricPressure = BarometricPressure
        self.WindSpeed = WindSpeed
        self.WindDirection = WindDirection
        self.RainIndicator = RainIndicator
        self.SnowIndicator = SnowIndicator
        self.DaylightSavingTimeIndicator = DaylightSavingTimeIndicator
        self.SolarModelIndicator = SolarModelIndicator
        self.BeamSolarDayScheduleName = BeamSolarDayScheduleName
        self.DiffuseSolarDayScheduleName = DiffuseSolarDayScheduleName
        self.ASHRAEClearSkyOpticalForBeamIrradiance = ASHRAEClearSkyOpticalForBeamIrradiance
        self.ASHRAEClearSkyOpticalForBDiffuseIrradiance = ASHRAEClearSkyOpticalForBDiffuseIrradiance
        self.SkyClearness = SkyClearness
        self.MaximumNumberWarmUpDays = MaximumNumberWarmUpDays
        self.BeginEnvironmentResetMode = BeginEnvironmentResetMode

class RunPeriod:
    def __init__(self,ClassName, Name, BeginMonth, BeginDayOfMonth, BeginYear, EndMonth, EndDayOfMonth, EndYear, DayOfWeekForStartDay, UseWeatherFileHolidaysAndSpecialDays, UseWeatherFileDaylightSavingPeriod, ApplyWeekendHolidayRule, UseWeatherFileRainIndicators, \
        UseWeatherFileSnowIndicators, TreatWeatherAsActual):
        self.ClassName = ClassName
        self.Name = Name
        self.BeginMonth = BeginMonth
        self.BeginDayOfMonth = BeginDayOfMonth
        self.BeginYear = BeginYear
        self.EndMonth = EndMonth
        self.EndDayOfMonth = EndDayOfMonth
        self.EndYear = EndYear
        self.DayOfWeekForStartDay = DayOfWeekForStartDay
        self.UseWeatherFileHolidaysAndSpecialDays = UseWeatherFileHolidaysAndSpecialDays
        self.UseWeatherFileDaylightSavingPeriod = UseWeatherFileDaylightSavingPeriod
        self.ApplyWeekendHolidayRule = ApplyWeekendHolidayRule
        self.UseWeatherFileRainIndicators = UseWeatherFileRainIndicators
        self.UseWeatherFileSnowIndicators = UseWeatherFileSnowIndicators
        self.TreatWeatherAsActual = TreatWeatherAsActual

class ScheduleTypeLimits:
    def __init__(self,ClassName, Name, LowerLimitValue, UpperLimitValue, NumericType, UnitType):
        self.ClassName = ClassName
        self.Name = Name
        self.LowerLimitValue = LowerLimitValue
        self.UpperLimitValue = UpperLimitValue
        self.NumericType = NumericType
        self.UnitType = UnitType

class ScheduleConstant:
    def __init__(self,ClassName, Name, ScheduleTypeLimitsName, HourlyValue):
        self.ClassName = ClassName
        self.Name = Name
        self.ScheduleTypeLimitsName = ScheduleTypeLimitsName
        self.HourlyValue = HourlyValue

class GlobalGeometryRules:
    def __init__(self,ClassName, StartingVertexPosition, VertexEntryDirection, CoordinateSystem, DayLightiningReferencePointCoordinateSystem, RectangularSurfaceCoordinateSystem):
        self.ClassName = ClassName
        self.StartingVertexPosition = StartingVertexPosition
        self.VertexEntryDirection = VertexEntryDirection
        self.CoordinateSystem = CoordinateSystem
        self.DayLightiningReferencePointCoordinateSystem = DayLightiningReferencePointCoordinateSystem
        self.RectangularSurfaceCoordinateSystem = RectangularSurfaceCoordinateSystem

class Zone:
    def __init__(self,ClassName, Name, DirectionOfRelativeNorth, XOrigin, YOrigin, ZOrigin, Type, Multiplier, CeilingHeight, Volume, FloorArea, ZoneInsideConvectionAlgorithm, ZoneOutsideConvectionAlgorithm, PartOfTotalFloorArea):
        self.ClassName = ClassName
        self.Name = Name
        self.DirectionOfRelativeNorth = DirectionOfRelativeNorth
        self.XOrigin = XOrigin
        self.YOrigin = YOrigin
        self.ZOrigin = ZOrigin
        self.Type = Type
        self.Multiplier = Multiplier
        self.CeilingHeight = CeilingHeight     
        self.Volume = Volume 
        self.FloorArea = FloorArea
        self.ZoneInsideConvectionAlgorithm = ZoneInsideConvectionAlgorithm 
        self.ZoneOutsideConvectionAlgorithm = ZoneOutsideConvectionAlgorithm
        self.PartOfTotalFloorArea = PartOfTotalFloorArea      

class BuildingSurface_Detailed:
    def __init__(self, ClassName, Name, SurfaceType, ConstructionName, ZoneName, OutsideBoundaryCondition, OutsideBoundaryConditionObject, SunExposure, WindExposure, ViewFactorToGround, NumberOfVertices):
        self.ClassName = ClassName
        self.Name = Name
        self.SurfaceType = SurfaceType
        self.ConstructionName = ConstructionName
        self.ZoneName = ZoneName
        self.OutsideBoundaryCondition = OutsideBoundaryCondition
        self.OutsideBoundaryConditionObject = OutsideBoundaryConditionObject
        self.SunExposure = SunExposure
        self.WindExposure = WindExposure     
        self.ViewFactorToGround = ViewFactorToGround
        self.NumberOfVertices = NumberOfVertices

class Construction:
    def __init__(self, ClassName, Name, OutsideLayer):
        self.ClassName = ClassName
        self.Name = Name
        self.OutsideLayer = OutsideLayer

class Material:
    def __init__(self, ClassName, Name, Roughness, Thickness, Conductivity, Density, SpecificHeat, ThermalAbsorptance, SolarAbsorptance, VisibleAbsorptance):
        self.ClassName = ClassName
        self.Name = Name
        self.Roughness = Roughness
        self.Thickness = Thickness
        self.Conductivity = Conductivity
        self.Density = Density
        self.SpecificHeat = SpecificHeat
        self.ThermalAbsorptance = ThermalAbsorptance
        self.SolarAbsorptance = SolarAbsorptance
        self.VisibleAbsorptance = VisibleAbsorptance

class Material_NoMass:
    def __init__(self, ClassName, Name, Roughness, ThermalResistance, ThermalAbsorptance, SolarAbsorptance, VisibleAbsorptance):
        self.ClassName = ClassName
        self.Name = Name
        self.Roughness = Roughness
        self.ThermalResistance = ThermalResistance
        self.ThermalAbsorptance = ThermalAbsorptance
        self.SolarAbsorptance = SolarAbsorptance
        self.VisibleAbsorptance = VisibleAbsorptance

class WindowMaterial_SimpleGlazingSystem:
    def __init__(self, ClassName, Name, UFactor, SolarHeatGainCoefficient, VisibleTransmitance):
        self.ClassName = ClassName
        self.Name = Name
        self.UFactor = UFactor
        self.SolarHeatGainCoefficient = SolarHeatGainCoefficient
        self.TVisibleTransmitance = VisibleTransmitance

class FenestrationSurface:
    def __init__(self, ClassName, Name, SurfaceType, ConstructionName, BuildingSurfaceName, OutsideBoundaryConditionObject, ViewFactorToTheGround,FrameAndDividerName, Multiplier, NumberOfVertices):
        self.ClassName = ClassName
        self.Name = Name
        self.SurfaceType = SurfaceType
        self.ConstructionName = ConstructionName
        self.BuildingSurfaceName = BuildingSurfaceName
        self.OutsideBoundaryConditionObject = OutsideBoundaryConditionObject
        self.ViewFactorToTheGround = ViewFactorToTheGround
        self.FrameAndDividerName = FrameAndDividerName
        self.Multiplier = Multiplier
        self.NumberOfVertices = NumberOfVertices

class Output_VariableDictionary:
    def __init__(self, ClassName, KeyField, SortOption):
        self.ClassName = ClassName
        self.Name = KeyField
        self.SurfaceType = SortOption

class Output_Surfaces_Drawing:
    def __init__(self, ClassName, ReportType, ReportSpecifications1, ReportSpecifications2):
        self.ClassName = ClassName
        self.ReportType = ReportType
        self.ReportSpecifications1 = ReportSpecifications1
        self.ReportSpecifications2 = ReportSpecifications2

class Output_Constructions:
    def __init__(self, ClassName, DetailsType1, DetailsType2):
        self.ClassName = ClassName
        self.DetailsType1 = DetailsType1
        self.DetailsType2 = DetailsType2

class Output_Table_SummaryReports:
    def __init__(self, ClassName, Report1Name):
        self.ClassName = ClassName
        self.Report1Name = Report1Name

class Output_Control_TableStyle:
    def __init__(self, ClassName, ColumnSeparator, UnitConversion):
        self.ClassName = ClassName
        self.ColumnSeparator = ColumnSeparator
        self.UnitConversion = UnitConversion

class Output_Variable:
    def __init__(self, ClassName, KeyValue, VariableName,ReportingFrequency,ScheduleName):
        self.ClassName = ClassName
        self.KeyValue = KeyValue
        self.VariableName = VariableName
        self.ReportingFrequency = ReportingFrequency
        self.ScheduleName = ScheduleName

class Output_Meter_MeterFileOnly:
    def __init__(self, ClassName, KeyName, ReportingFrequency):
        self.ClassName = ClassName
        self.KeyName = KeyName
        self.ReportingFrequency = ReportingFrequency



#creating an error list to be presented to the user at the end of the execution

errors_list = []

#defining view and transaction for creating a energy analysis model
v = doc.ActiveView
t = Transaction(doc, "energy analysis model")


#creating a energy analysis model in Revit
t.Start()
settings = EnergyDataSettings.GetFromDocument(doc)
settings.AnalysisType = AnalysisMode.ConceptualMassesAndBuildingElements
options = EnergyAnalysisDetailModelOptions()
#include constructions, schedules, and non-graphical data in the computation of the energy analysis model
options.Tier = EnergyAnalysisDetailModelTier.Final
#energy analysis based on room and spaces
#BuildingElement
#SpatialElement
options.EnergyModelType = EnergyModelType.SpatialElement
#Create a new energy analysis detailed model from the physical model
eadm = EnergyAnalysisDetailModel.Create(doc, options) 
t.Commit()



spaces_list = eadm.GetAnalyticalSpaces()
surfaces_list = eadm.GetAnalyticalSurfaces()
openings_list = eadm.GetAnalyticalOpenings()


#Creating instances of the declared classes-------------------------------------------------------------------------------------------------------------------------------

#Creating an instance list
ObjectInstances = []

#Creating a Version instance
IDFVersion = Version("Version",str(9.4))
ObjectInstances.append(IDFVersion)

#Creating a SimulationCOntrol instance__________________________________________________________________________________________________________________
IDFSimulationControl = SimulationControl("SimulationControl",'No','No', 'No', 'Yes', 'Yes', 'No', 1)
ObjectInstances.append(IDFSimulationControl)


#Creating a Bulding instance____________________________________________________________________________________________________________________________

#getting the DirectionOfRelativeNorth (EP) from the Angle to True North (Revit)
origin = XYZ(0.0,0.0,0.0)
AngleToTrueNorth = UnitUtils.ConvertFromInternalUnits(doc.ActiveProjectLocation.GetProjectPosition(origin).Angle, UnitTypeId.Degrees)

IDFBuilding = Building("Building", doc.ProjectInformation.BuildingName, AngleToTrueNorth*-1 , 'Suburbs', str(0.05), str(0.05), 'MinimalShadowing',str(30),str(6)) 
ObjectInstances.append(IDFBuilding)

#Creating SurfaceConvectionAlgorithm instances__________________________________________________________________________________________________________
#Inside instance
IDFSurfaceConvectionAlgorithm = SurfaceConvectionAlgorithm("SurfaceConvectionAlgorithm:Inside", 'TARP') 
ObjectInstances.append(IDFSurfaceConvectionAlgorithm)
#Outside instance
IDFSurfaceConvectionAlgorithm = SurfaceConvectionAlgorithm("SurfaceConvectionAlgorithm:Outside", 'DOE-2') 
ObjectInstances.append(IDFSurfaceConvectionAlgorithm)

#Creating HeatBalanceAlgorithm instances________________________________________________________________________________________________________________

IDFHeatBalanceAlgorith = HeatBalanceAlgorithm("HeatBalanceAlgorithm", "ConductionTransferFunction", "","","") 
ObjectInstances.append(IDFHeatBalanceAlgorith)

#Creating Timestep instances____________________________________________________________________________________________________________________________

IDFTimestep = Timestep("Timestep", 4) 
ObjectInstances.append(IDFTimestep)

#Creating Site:Location instances_______________________________________________________________________________________________________________________


ProjectPlaceName =  doc.SiteLocation.PlaceName
ProjectLatitude =  doc.SiteLocation.Latitude
ProjectLongitude =  doc.SiteLocation.Longitude
ProjectTimeZone =  doc.SiteLocation.TimeZone
ProjectElevation =  doc.SiteLocation.Elevation

IDFSite = Site("Site:Location", ProjectPlaceName.replace(",", " "), ProjectLatitude, ProjectLongitude, ProjectTimeZone, ProjectElevation) 
ObjectInstances.append(IDFSite)

#Creating SizePeriod:DesignDay instances________________________________________________________________________________________________________________

#attributes filled manually
IDFSizingPeriod = SizingPeriod('SizingPeriod:DesignDay', 'Denver Centennial Winter', 12, 21, 'WinterDesignDay', -15.5, 0 , '', '', 'Wetbulb', -15.5, '', '', \
'', '', 81198, 3, 340, 'No', 'No', 'No', 'ASHRAEClearSky', '', '', \
'', '', 0, '', '')
ObjectInstances.append(IDFSizingPeriod)

IDFSizingPeriod = SizingPeriod('SizingPeriod:DesignDay', 'Denver Centennial Summer', 7, 21, 'SummerDesignDay', 32, 15.2 , '', '', 'Wetbulb', 15.5, '', '', \
'', '', 81198, 4.9, 0, 'No', 'No', 'No', 'ASHRAEClearSky', '', '', \
'', '', 0, '', '')
ObjectInstances.append(IDFSizingPeriod)

#Creating RunPeriod instances___________________________________________________________________________________________________________________________

#attributes filled manually
IDFRunPeriod= RunPeriod('RunPeriod', 'Run Period 1', 1, 1, '', 12, 31, '', 'Tuesday', 'Yes','Yes', 'No', 'Yes', \
        'Yes', '') 
ObjectInstances.append(IDFRunPeriod)

#Creating ScheduleTypeLimits instances__________________________________________________________________________________________________________________
#attributes filled manually

IDFScheduleTypeLimits= ScheduleTypeLimits('ScheduleTypeLimits', 'Fraction', 0, 1, 'CONTINUOUS', '') 
ObjectInstances.append(IDFScheduleTypeLimits)

IDFScheduleTypeLimits= ScheduleTypeLimits('ScheduleTypeLimits', 'On/Off', 0, 1, 'DISCRETE', '') 
ObjectInstances.append(IDFScheduleTypeLimits)

#Creating ScheduleConstant instances____________________________________________________________________________________________________________________

#attributes filled manually
IDFScheduleConstant= ScheduleConstant('Schedule:Constant', 'AlwaysOn', 'On/Off', 1) 
ObjectInstances.append(IDFScheduleConstant)

#Creating GlobalGeometryRules instances____________________________________________________________________________________________________________________

#attributes filled manually
IDFGlobalGeometryRules= GlobalGeometryRules('GlobalGeometryRules', 'LowerLeftCorner', 'Clockwise', 'Relative', '', '') 
ObjectInstances.append(IDFGlobalGeometryRules)

#Creating Zone instances________________________________________________________________________________________________________________________________

for space in spaces_list:
    #getting the space origin (using the first surface of the surface_list and the first point of the list of points of the selected surface)
    space_origin_X = round(UnitUtils.ConvertFromInternalUnits(surfaces_list[0].GetPolyloop().GetPoints()[0].X, UnitTypeId.Meters),6)
    space_origin_Y = round(UnitUtils.ConvertFromInternalUnits(surfaces_list[0].GetPolyloop().GetPoints()[0].Y, UnitTypeId.Meters),6)
    space_origin_Z = round(UnitUtils.ConvertFromInternalUnits(surfaces_list[0].GetPolyloop().GetPoints()[0].Z, UnitTypeId.Meters),6)
    #creating the zone instance
    #IDFZone = Zone("Zone",space.ComposedName, str(AngleToTrueNorth), space_origin_X, space_origin_Y, space_origin_Z, 1, 1, "autocalculate", "autocalculate", "", "", "", "")
    IDFZone = Zone("Zone",space.ComposedName, 0, space_origin_X, space_origin_Y, space_origin_Z, 1, 1, "autocalculate", "autocalculate", "", "", "", "")
    ObjectInstances.append(IDFZone)

#CREATING SURFACES INSTANCES

construction_list= []
opening_surface_name_list = []
link_opening_surface = []
surface_index = 0
for surface in surfaces_list:
    surface_index = surface_index + 1
    #Mapping the EnergyPlus properties OutisideBoundaryConditions, SunExposure and WindExposure
    OutsideBoundaryCondition = ""
    SunExposure = ""
    WindExposure = ""
    SurfaceType = ""
    if str(surface.Type) == "ExteriorWall":
        SurfaceType ="Wall" 
        OutsideBoundaryCondition = "Outdoors"
        SunExposure = "SunExposed"
        WindExposure = "WindExposed"
    if str(surface.Type) == "InteriorWall":
        SurfaceType ="Wall" 
        OutsideBoundaryCondition = "Outdoors"
        SunExposure = "NoSun"
        WindExposure = "NoWind"
    if str(surface.Type) == "InteriorFloor":
        SurfaceType ="Floor" 
        OutsideBoundaryCondition = "Outdoors"
        SunExposure = "SunExposed"
        WindExposure = "WindExposed"
    if str(surface.Type) == "Ceiling":
        SurfaceType ="Ceiling" 
        OutsideBoundaryCondition = "Outdoors"
        SunExposure = "SunExposed"
        WindExposure = "WindExposed"
    if str(surface.Type) == "SlabOnGrade":
        OutsideBoundaryCondition = "Adiabatic"
        SunExposure = "NoSun"
        WindExposure = "NoWind"

    surface_construction_name = surface.OriginatingElementDescription
    #eliminating special caracters not acceptable on EnergyPlus
    string_index = surface_construction_name.find("[")
    surface_construction_name = surface_construction_name[0:string_index-1:]

    #getting the object that generated the surface
    surface_element = doc.GetElement(str(surface.CADObjectUniqueId))

    #getting the surface element that hosts the opnening
    surface_analytical_openings = surface.GetAnalyticalOpenings()
    for analytical_opening in surface_analytical_openings:
        link_opening_surface.append([analytical_opening.UniqueId,"Obj" + str(surface_index) + ": ", surface_element.UniqueId])

    IDFSurface = BuildingSurface_Detailed("BuildingSurface:Detailed","Obj" + str(surface_index) + ": " + surface_element.UniqueId , SurfaceType , surface_construction_name, surface.GetAnalyticalSpace().ComposedName, OutsideBoundaryCondition,"", SunExposure, WindExposure, "", len(surface.GetPolyloop().GetPoints()))
    for i in range(len(surface.GetPolyloop().GetPoints())):
        setattr(IDFSurface, "Vertex " + str(i +1) + " X-coordinate", round(UnitUtils.ConvertFromInternalUnits(surface.GetPolyloop().GetPoints()[i].X, UnitTypeId.Meters),6))
        setattr(IDFSurface, "Vertex " + str(i +1) + " Y-coordinate", round(UnitUtils.ConvertFromInternalUnits(surface.GetPolyloop().GetPoints()[i].Y, UnitTypeId.Meters),6))
        setattr(IDFSurface, "Vertex " + str(i +1) + " Z-coordinate", round(UnitUtils.ConvertFromInternalUnits(surface.GetPolyloop().GetPoints()[i].Z, UnitTypeId.Meters),6))
    ObjectInstances.append(IDFSurface)


    #Creating a list of the GUID of the elements that originated the AnalysisSurfaces through the property CADObjectUniqueId
    for i in range(len(construction_list)):
        if i < len(construction_list):    
            if construction_list[i] == surface.CADObjectUniqueId:
                construction_list.remove(construction_list[i])
    construction_list.append(surface.CADObjectUniqueId)
    for construction in construction_list:
        construction_element = doc.GetElement(str(construction))
        construction_element_type = doc.GetElement(construction_element.GetTypeId())
        construction_revit_type = construction_element_type.GetParameters("Type Name")[0].AsString()
        if construction_revit_type == "100 POR 100":
            print construction_element.Id

    #Creating a list of the GUID of the elements that originated the AnalysisSurfaces through the property CADObjectUniqueId + surface name
    for i in range(len(opening_surface_name_list)):
        if i < len(opening_surface_name_list):    
            if opening_surface_name_list[i] == surface.CADObjectUniqueId:
                opening_surface_name_list.remove(opening_surface_name_list[i])
    opening_surface_name_list.append([surface.CADObjectUniqueId, surface.GetAnalyticalSpace().ComposedName +": " + surface.Name])



# Creating openings (fenestration) instances____________________________________________________________________________________________________________

#eliminating any duplicated opening on the opening list
auxiliar_openings_list = []
Id_openings_list = []
n = 0
for i in range(len(openings_list)):
    auxiliar_openings_list.append([openings_list[i],doc.GetElement(str(openings_list[i].CADObjectUniqueId)).UniqueId])
    Id_openings_list.append(doc.GetElement(str(openings_list[i].CADObjectUniqueId)).UniqueId)

Id_openings_list = list(dict.fromkeys(Id_openings_list))
for i in range(len(Id_openings_list)):
    n = 0
    for auxiliar_opening in auxiliar_openings_list:
        if Id_openings_list[i] == auxiliar_opening[1]:
            n = n +1
            if n > 1:
                auxiliar_openings_list.remove(auxiliar_opening)

openings_list = []
for auxiliar_opening in auxiliar_openings_list:
    openings_list.append(auxiliar_opening[0])

for opening in openings_list:
    #getting opening type (window, door, dome, etc)
    opening_surface_type = str(opening.OpeningType)
    #getting the object that generated the opening
    opening_element = doc.GetElement(str(opening.CADObjectUniqueId))
    #getting the surface that host the object that generated the opening
    for link_opening_surface_item in link_opening_surface:
        if opening.UniqueId == link_opening_surface_item[0]:
            surface_id = link_opening_surface_item[1] + link_opening_surface_item [2]
            break
    for i in range(len(opening_surface_name_list)):
        if i < len(opening_surface_name_list):
            if str(opening_surface_name_list[i][0]) == surface_id:
                opening_surface_name = opening_surface_name_list[i][1]
    #getting the opening construction name
    element_type = doc.GetElement(opening_element.GetTypeId())
    family_name = element_type.FamilyName
    revit_type = element_type.GetParameters("Type Name")[0].AsString()
    if element_type.GetParameters("Analytic Construction")[0].AsString() != None:	
        opening_construction_name = element_type.FamilyName + ": " + element_type.GetParameters("Analytic Construction")[0].AsString()
    else:
        opening_construction_name = str(family_name) + ": " + str(revit_type)
    #creating opening instances and adding the vertexes according to the opening geometry
    #opening surface type
    #IDFFenestrationSurface = FenestrationSurface("FenestrationSurface:Detailed", opening_surface_name +"_" + opening_surface_type , opening_surface_type , opening_construction_name, opening_surface_name, "","autocalculate", "", 1 , len(opening.GetPolyloop().GetPoints()))
    IDFFenestrationSurface = FenestrationSurface("FenestrationSurface:Detailed", opening_element.UniqueId , 'Window' , opening_construction_name, surface_id, "","autocalculate", "", 1 , len(opening.GetPolyloop().GetPoints()))
    for i in range(len(opening.GetPolyloop().GetPoints())):
        setattr(IDFFenestrationSurface, "Vertex " + str(i +1) + " X-coordinate", round(UnitUtils.ConvertFromInternalUnits(opening.GetPolyloop().GetPoints()[i].X, UnitTypeId.Meters),6))
        setattr(IDFFenestrationSurface, "Vertex " + str(i +1) + " Y-coordinate", round(UnitUtils.ConvertFromInternalUnits(opening.GetPolyloop().GetPoints()[i].Y, UnitTypeId.Meters),6))
        setattr(IDFFenestrationSurface, "Vertex " + str(i +1) + " Z-coordinate", round(UnitUtils.ConvertFromInternalUnits(opening.GetPolyloop().GetPoints()[i].Z, UnitTypeId.Meters),6))
    ObjectInstances.append(IDFFenestrationSurface)

    for i in range(len(construction_list)):
        if i < len(construction_list):    
            if construction_list[i] == opening.CADObjectUniqueId:
                construction_list.remove(construction_list[i])
    construction_list.append(opening.CADObjectUniqueId)



#Creating Construction class instances from the building surfaces________________________________________________________________________    
materials_list= []
materialsnomass_list= []
materials_windows= []

#eliminating duplicated elements from the construction list

auxiliar_list = []
non_duplicated_construction_list = []

for construction in construction_list:
    construction_element = doc.GetElement(str(construction))
    construction_element_type = doc.GetElement(construction_element.GetTypeId())
    construction_revit_type = construction_element_type.GetParameters("Type Name")[0].AsString()	
    non_duplicated_construction_list.append(construction_revit_type)

#Ordering the list alphabetically 
non_duplicated_construction_list = sorted(non_duplicated_construction_list, key=operator.itemgetter(0))
#Removing duplicates from the list
non_duplicated_construction_list = list(dict.fromkeys(non_duplicated_construction_list))

for non_duplicated in non_duplicated_construction_list:
    for construction in construction_list:
        construction_element = doc.GetElement(str(construction))
        construction_element_type = doc.GetElement(construction_element.GetTypeId())
        construction_revit_type = construction_element_type.GetParameters("Type Name")[0].AsString()
        if non_duplicated == construction_revit_type:
            auxiliar_list.append (construction)
            break
        
construction_list = auxiliar_list


for construction in construction_list:
    construction_element = doc.GetElement(str(construction))
    construction_element_type = doc.GetElement(construction_element.GetTypeId())
    construction_revit_type = construction_element_type.GetParameters("Type Name")[0].AsString()
    

for construction in construction_list:
    #getting the originating revit element based on the GUID
    construction_element = doc.GetElement(str(construction))
    #getting the type of the object
    element_type = doc.GetElement(construction_element.GetTypeId())
    family_name = element_type.FamilyName
    revit_type = element_type.GetParameters("Type Name")[0].AsString()	
    construction_name = str(family_name) + ": " + str(revit_type)
    if str(element_type.Category.Name) == "Walls" or str(element_type.Category.Name) == "Floors" or str(element_type.Category.Name) == "Roofs" or str(element_type.Category.Name) == "Ceilings":
        compound_structure_list = element_type.GetCompoundStructure().GetLayers()
        print compound_structure_list
        #FALTA ADICIONAR MAIS LAYERS SER O OBJETO TIVER MAIS DE UMA
        #Reogarnizing list from the most exterior Layer0 to the most interior layer
        compound_structure_reordered_list = []
        for compound_structure in compound_structure_list:
            for i in range(len(compound_structure_reordered_list)):
                if i < len(compound_structure_reordered_list):    
                    if compound_structure_reordered_list[i] == [compound_structure.MaterialId, compound_structure.LayerId]:
                        compound_structure_reordered_list.remove(compound_structure_reordered_list[i])
            #converting the compound layer thickness to meters
            thickness = UnitUtils.ConvertFromInternalUnits(compound_structure.Width, UnitTypeId.Meters)
            compound_structure_reordered_list.append([compound_structure.MaterialId,compound_structure.LayerId, thickness])
            
            #classifying the element roughness according to Revit and EnergyPlus (material properties)
            if element_type.LookupParameter('Roughness') != None:
                material_roughness_number = element_type.ThermalProperties.Roughness
                if material_roughness_number == 1:
                    material_roughness = "VeryRough"
                if material_roughness_number == 2:
                    material_roughness = "Rough"
                if material_roughness_number == 3:
                    material_roughness = "MediumRough"
                if material_roughness_number == 4:
                    material_roughness = "MediumSmooth"
                if material_roughness_number == 5:
                    material_roughness = "Smooth"
                if material_roughness_number == 6:
                    material_roughness = "VerySmooth"

            for i in range(len(materials_list)):
                if i < len(materials_list):    
                    if materials_list[i] == [compound_structure.MaterialId, thickness, material_roughness] :
                        materials_list.remove(materials_list[i])
            materials_list.append([compound_structure.MaterialId, thickness, material_roughness])

            #registring inconsistences in the error_list on the objects that do not hava a material assigned or has no thickness
            if str(compound_structure.MaterialId) == str(-1):
                errors_list.append("The object [" + revit_type + "] has no materials assigned for some of its compound layers")
            if thickness == 0 or thickness == None:
                errors_list.append("The object [" + revit_type + "] has no thickness assigned for some of its compound layers. A minimal value of 0.001m was assumed")
        compound_structure_reordered_list = sorted(compound_structure_reordered_list, key=operator.itemgetter(1,0))
        if doc.GetElement(compound_structure_reordered_list[0][0]) == None:
            material_externo = "NoMaterial"
        else:
            material_externo = doc.GetElement(compound_structure_reordered_list[0][0]).Name
        IDFConstruction = Construction("Construction", construction_name, material_externo.replace(",", " ") + " " + str(compound_structure_reordered_list[0][2]) + " m" )
        for i in range(len(compound_structure_reordered_list)):
            if i > 0:
                setattr(IDFConstruction, "Layer " + str(i +1), doc.GetElement(compound_structure_reordered_list[i][0]).Name.replace(",", " ") + " " + str(compound_structure_reordered_list[i][2]) + " m")
        ObjectInstances.append(IDFConstruction)
    else:
        #if the construction instances was generated by an opening, the following code uses the analytic construction Revit parameter to be used as construction name
        if element_type.Category.Name =="Doors" or element_type.Category.Name =="Windows":
            if element_type.GetParameters("Analytic Construction")[0].AsString() != None:
                construction_name = element_type.GetParameters("Analytic Construction")[0].AsString()
            else:
                construction_name = str(revit_type)
                #registring inconsistences in the error_list on the objects that do not hava a material assigned or has no thickness
                errors_list.append("The object [" + revit_type + "] has no analytical construction assigned, therefore, the thermal properties are not filled")
            #if the construction instances was generated by an opening, the following code gets the Thermal parameters to fill the classes MaterialWindows
            materials_window_name = element_type.GetParameters("Analytic Construction")[0].AsString()
            HeatTransferCoefficientU = element_type.GetParameters("Heat Transfer Coefficient (U)")[0].AsDouble()
            VisualLightTransmittance = element_type.GetParameters("Visual Light Transmittance")[0].AsDouble()
            if VisualLightTransmittance == 0:
                VisualLightTransmittance = 0.000001
                errors_list.append("The object [" + revit_type + "] has no Visual Light Transmittance assigned. A minimal value of 0.000001 was assumed")
            ThermalResistanceR = element_type.GetParameters("Thermal Resistance (R)")[0].AsDouble()	
            SolarHeatGainCoefficient = element_type.GetParameters("Solar Heat Gain Coefficient")[0].AsDouble()
            if SolarHeatGainCoefficient == 0:
                SolarHeatGainCoefficient = 0.000001
                errors_list.append("The object [" + revit_type + "] has no Visual Solar Heat Gain Coefficient assigned. A minimal value of 0.000001 was assumed")
            for i in range(len(materials_windows)):
                    if i < len(materials_windows):    
                        if materials_windows[i] == [materials_window_name, HeatTransferCoefficientU, VisualLightTransmittance,ThermalResistanceR,SolarHeatGainCoefficient] :
                            materials_windows.remove(materials_windows[i])
            materials_windows.append([materials_window_name, HeatTransferCoefficientU, VisualLightTransmittance,ThermalResistanceR,SolarHeatGainCoefficient])
            #creating a construction IDF instance 
            IDFConstruction = Construction("Construction", element_type.FamilyName + ": " + construction_name, materials_window_name)
            ObjectInstances.append(IDFConstruction)


# Creating Material and No Mass materials instances based on the density
for element in materials_list:
    material = doc.GetElement(element[0])
    if material == None:
        IDFMaterialNoMass = Material_NoMass("Material:NoMass", "Material not assigned", "", "", "", "", "")
        ObjectInstances.append(IDFMaterialNoMass)
    else:
        material_name = material.Name
        thermal_asset = doc.GetElement(material.ThermalAssetId).GetThermalAsset()
        #calculating the thermal resistance R = thickness / U
        if thermal_asset.ThermalConductivity != 0:
            #assuming 0,001 thickness [element[01]] for the elements with no thickness specified, and assuming 0.001 thermal resistance for layers with thermal resistance inferior to 0.001
            if element[1] == 0 or element[1] == None:
                material_ThermalResistance = 0.001 / thermal_asset.ThermalConductivity
                if material_ThermalResistance < 0.001:
                    material_ThermalResistance = 0.001
            else:  
                material_ThermalResistance = element[1] / thermal_asset.ThermalConductivity
        # element [2] is material_roughness
        if thermal_asset.Density == None or element[1] == 0:
            IDFMaterialNoMass = Material_NoMass("Material:NoMass", material_name.replace(",", " "), element[2], material_ThermalResistance, thermal_asset.Emissivity, "", "")
            ObjectInstances.append(IDFMaterialNoMass)
        else:
            IDFMaterial = Material("Material", material_name.replace(",", " ") +" " + str(element[1]) + " m", element[2], element[1], thermal_asset.ThermalConductivity, thermal_asset.Density, thermal_asset.SpecificHeat, thermal_asset.Emissivity, "", "")
            ObjectInstances.append(IDFMaterial)

#Creating WindowMaterial:SimpleGlazingSystem class instances________________________________________________________________________    

for window_material in materials_windows:
    IDFMaterial = WindowMaterial_SimpleGlazingSystem("WindowMaterial:SimpleGlazingSystem", window_material[0], window_material[1], window_material[4],window_material[2])
    ObjectInstances.append(IDFMaterial)

#Creating a Output:VariableDictionary instance__________________________________________________________________________________________________________________
IDFSimulationControl = Output_VariableDictionary("Output:VariableDictionary",'IDF',"")
ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Surfaces:Drawing instance__________________________________________________________________________________________________________________
IDFSimulationControl = Output_Surfaces_Drawing("Output:Surfaces:Drawing",'dxf:wireframe',"","")
ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Surfaces:Drawing instance__________________________________________________________________________________________________________________
IDFSimulationControl = Output_Constructions("Output:Constructions",'Constructions',"")
ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Table:SummaryReports instance__________________________________________________________________________________________________________________
IDFSimulationControl = Output_Table_SummaryReports("Output:Table:SummaryReports", 'AllSummary')
ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Table:SummaryReports instance__________________________________________________________________________________________________________________
IDFSimulationControl = Output_Control_TableStyle("OutputControl:Table:Style", 'ALL', "")
ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Variable instance__________________________________________________________________________________________________________________

output_variable_list = [["*",'Site Outdoor Air Drybulb Temperature','hourly'],["*",'Site Daylight Saving Time Status','daily'],["*",'Site Day Type Index','daily'],
["*",'Zone Mean Air Temperature','hourly'],["*",'Zone Total Internal Latent Gain Energy','hourly'],["*",'Zone Mean Radiant Temperature','hourly'],
["*",'Zone Air Heat Balance Surface Convection Rate','hourly'],["*",'Zone Air Heat Balance Air Energy Storage Rate','hourly'],["*",'Surface Inside Face Temperature','daily'],
["*",'Surface Outside Face Temperature','daily'],["*",'Surface Inside Face Convection Heat Transfer Coefficient','daily'],["*",'Surface Outside Face Convection Heat Transfer Coefficient','daily'],
["*",'Other Equipment Total Heating Energy','monthly'],["*",'Zone Other Equipment Total Heating Energy','monthly']]

for output_variable in output_variable_list:
    IDFSimulationControl = Output_Variable("Output:Variable", output_variable[0], output_variable[1],output_variable[2],'')
    ObjectInstances.append(IDFSimulationControl)

#Creating a Output:Meter:MeterFileOnly instance__________________________________________________________________________________________________________________

IDFSimulationControl = Output_Meter_MeterFileOnly("Output:Meter:MeterFileOnly", 'ExteriorLights:Electricity', 'hourly')
ObjectInstances.append(IDFSimulationControl)
IDFSimulationControl = Output_Meter_MeterFileOnly("Output:Meter:MeterFileOnly", 'EnergyTransfer:Building', 'hourly')
ObjectInstances.append(IDFSimulationControl)
IDFSimulationControl = Output_Meter_MeterFileOnly("Output:Meter:MeterFileOnly", 'EnergyTransfer:Facility', 'hourly')
ObjectInstances.append(IDFSimulationControl)

directory = 'C:\Users\lucas\Documents'
filename = "teste31.idf"
file_path = os.path.join(directory, filename)
if not os.path.isdir(directory):
    os.mkdir(directory)
file = open(file_path, "w")
#writing each instance of object in the IDF file | ObjectInstances: list of all instances to be printed
for instance in ObjectInstances:
    file.write(str(instance.ClassName + ",\n"))
    #writing each attribute of the instance in the IDF file: 
    #instance.__dict__.keys returns a list of the attribute names contained in that instance
    #instance.__dict__.values returns a list with the values contained in the instance attributes
    for attribute in range(len(instance.__dict__.keys())):
        if attribute < len(instance.__dict__.keys()):
            if attribute == len(instance.__dict__.keys())-1:
                file.write("    " + str(instance.__dict__.values()[attribute]) + ";    !-" + str(instance.__dict__.keys()[attribute]) + "\n")
            else:
                if str(instance.__dict__.keys()[attribute]) != "ClassName":
                    file.write("    " + str(instance.__dict__.values()[attribute]) + ",    !-" + str(instance.__dict__.keys()[attribute]) + "\n")
file.close()

#removing duplicated items from the errors_list
errors_list = list(dict.fromkeys(errors_list))

message = ""

errors_counter = 0
if len(errors_list) < 1:
    ctypes.windll.user32.MessageBoxW(0, "The Energy Model (IDF) was successfully generated.", "IDF Exporter", 0)
else:
    for error in errors_list:
        errors_counter = errors_counter + 1
        message = message + "\n" + str(errors_counter) + "-" + error + "; "
    ctypes.windll.user32.MessageBoxW(0, "The Energy Model (IDF) was generated with the following inconsistences: " + message, "IDF Exporter", 0)