import os
import pathlib
import ezdxf
from object import *
from function_processingDXF import *
from function_processingExcel import *


if __name__=="__main__":

    #Step1. Get CAD file path
    str_CAD_file_name="example.dxf"
    pathInputFilePath=pathlib.Path(os.getcwd())
    pathInputFilePath=pathInputFilePath.joinpath(str_CAD_file_name)
    input_file_path = str(pathInputFilePath)

    #Step2. Open CAD file and model space
    doc = ezdxf.readfile(input_file_path)
    msp = doc.modelspace()

    #Step3. create floor data sheet record imformation of select area
    FloorDataSheet = []
    getFloorBorder(msp, "SelectArea", FloorDataSheet)
    getFloorName(msp, "floor_name", FloorDataSheet)
    getArea(msp, "41-樓地板面積", FloorDataSheet, "area_floor")
    getArea(msp, "42-梯廳面積", FloorDataSheet, "area_hall")
    getArea(msp, "43-管委會面積", FloorDataSheet, "area_committee")
    getArea(msp, "44-機電面積", FloorDataSheet, "area_electromechanical")
    getArea(msp, "45-車道面積", FloorDataSheet, "area_carramp")
    getArea(msp, "46-陽台面積", FloorDataSheet, "area_balcony")

    #Step4. Add floor information on the new dxf file. 
    createCalculateFormula(msp, FloorDataSheet)

    #Step5. Save processed file
    file_name_parts = input_file_path.split('.')
    output_file_path0 = file_name_parts[0] + "-after.dxf"
    doc.saveas(output_file_path0)

    #Step6. Save value in excel file
    createExcel(FloorDataSheet, "面積計算表")
