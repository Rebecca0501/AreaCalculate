from object import *
from ezdxf.enums import TextEntityAlignment

def getFloorBorder(msp, LayerName:str, FloorDataSheet):
    for LWPOLYLINE in msp.query("LWPOLYLINE[layer=='{}']".format(LayerName)):
        LWP_point = LWPOLYLINE.get_points()
        p1 = CPoint(LWP_point[0][0],LWP_point[0][1])
        p2 = CPoint(LWP_point[1][0],LWP_point[1][1])
        p3 = CPoint(LWP_point[2][0],LWP_point[2][1])
        p4 = CPoint(LWP_point[3][0],LWP_point[3][1])
        select_area = CSqure(CLine(p1,p2),CLine(p4,p3))
        
        x_min = select_area.x_min()
        x_max = select_area.x_max()
        y_min = select_area.y_min()
        y_max = select_area.y_max()
        new_floor = floor_imformation(x_min,x_max,y_min,y_max)
        FloorDataSheet.append(new_floor)

def getFloorName(msp, LayerName:str, FloorDataSheet):
    for TEXT in msp.query("TEXT[layer=='{}']".format(LayerName)):
        TextCoordinate = TEXT.dxf.insert
        TextCoordiPoint = CPoint(TextCoordinate[0],TextCoordinate[1])
        text_py = TEXT.dxf.text
        for idx, FloorInfo in enumerate(FloorDataSheet):
            if FloorInfo.x_min < TextCoordiPoint.x and TextCoordiPoint.x < FloorInfo.x_max and FloorInfo.y_min < TextCoordiPoint.y and TextCoordiPoint.y < FloorInfo.y_max:
                FloorDataSheet[idx].name = text_py
                break

    for MTEXT in msp.query("MTEXT[layer=='{}']".format(LayerName)):
        TextCoordinate = MTEXT.dxf.insert
        TextCoordiPoint = CPoint(TextCoordinate[0],TextCoordinate[1])
        text_py = MTEXT.dxf.text
        for idx, FloorInfo in enumerate(FloorDataSheet):
            if FloorInfo.x_min < TextCoordiPoint.x and TextCoordiPoint.x < FloorInfo.x_max and FloorInfo.y_min < TextCoordiPoint.y and TextCoordiPoint.y < FloorInfo.y_max:
                FloorDataSheet[idx].name = text_py
                break

def getArea(msp, LayerName:str, FloorDataSheet, attribute_name):
    for LWPOLYLINE in msp.query("LWPOLYLINE[layer=='{}']".format(LayerName)):
        points = point_list(LWPOLYLINE.get_points()).result()
        sumOfArea = 0.00
        IntPointNum = len(points)
        for idx, f in enumerate(FloorDataSheet):
            if f.x_min < points[0].x and points[0].x < f.x_max and f.y_min < points[0].y and points[0].y < f.y_max:
                #****************Be careful unit conversion**********
                sum_detA = abs(detA_sum(IntPointNum,points).result())*0.0001
                #****************Be careful unit conversion**********
                sumOfArea += sum_detA
                break
        match attribute_name:
            case "area_floor":
                FloorDataSheet[idx].area_floor = sumOfArea
            case "area_hall":
                FloorDataSheet[idx].area_hall = sumOfArea
            case "area_committee":
                FloorDataSheet[idx].area_committee = sumOfArea
            case "area_electromechanical":
                FloorDataSheet[idx].area_electromechanical = sumOfArea
            case "area_carramp":
                FloorDataSheet[idx].area_carramp = sumOfArea
            case "area_balcony":
                FloorDataSheet[idx].area_balcony = sumOfArea



def addTextLine(msp, FloorInfo, Content: str, line:int):
    msp.add_text(
        Content,
        height=FloorInfo.font_size,
    ).set_placement((FloorInfo.write_point.x, FloorInfo.write_point.y), align=TextEntityAlignment.LEFT)
    FloorInfo.write_point.y -= (FloorInfo.font_size*1.5*line)

def createCalculateFormula(msp, FloorDataSheet):
    for idx, FloorInfo in enumerate(FloorDataSheet):

        FloorInfo.area_floorRemoveCar = FloorInfo.area_floor - FloorInfo.area_carramp

        FloorInfo.over_hall =  max(FloorInfo.area_hall - (FloorInfo.area_floorRemoveCar*0.1),0)

        FloorInfo.over_balcony =  max(FloorInfo.area_balcony - (FloorInfo.area_floorRemoveCar*0.1),0)

        FloorInfo.over_hallBalcony = max(FloorInfo.area_hall + FloorInfo.area_balcony - ((FloorInfo.area_floorRemoveCar)*0.15) - FloorInfo.over_hall - FloorInfo.over_balcony,0)

        FloorInfo.area_volumn = FloorInfo.area_floor-FloorInfo.area_hall-FloorInfo.area_committee-FloorInfo.area_electromechanical-FloorInfo.area_carramp + FloorInfo.over_hall + FloorInfo.over_balcony + FloorInfo.over_hallBalcony

        if FloorInfo.area_floor >0:
            content = str(f"樓地板面積: {round(FloorInfo.area_floor,2)}")
            addTextLine(msp, FloorInfo, content, 2)

        if FloorInfo.area_carramp >0.:
            content = str(f"室內停車&車道面積: {round(FloorInfo.area_carramp,2)}")
            addTextLine(msp, FloorInfo, content, 2)

            content = str("扣除室內停車面積之樓地板面積:")
            addTextLine(msp, FloorInfo, content, 1)
            
            content = str(f"= {round(FloorInfo.area_floor,2)}(樓地板面積) - {round(FloorInfo.area_carramp,2)}(車道面積) = {round(FloorInfo.area_floorRemoveCar,2)}")
            addTextLine(msp, FloorInfo, content, 2)

        if FloorInfo.area_hall >0.:
            content = str(f"梯廳面積: {round(FloorInfo.area_hall,2)}")
            addTextLine(msp, FloorInfo, content, 1)

            if FloorInfo.over_hall<0.001:
                content = str(f"{round(FloorInfo.area_hall,2)} < {round(FloorInfo.area_floorRemoveCar,2)}*0.1 = {round(FloorInfo.area_floorRemoveCar*0.1,2)}...OK")
                addTextLine(msp, FloorInfo, content, 2)
            else:
                content = str(f"{round(FloorInfo.area_hall,2)} > {round(FloorInfo.area_floorRemoveCar,2)}*0.1 = {round(FloorInfo.area_floorRemoveCar*0.1,2)}")
                addTextLine(msp, FloorInfo, content, 1)

                content = str(f"{round(FloorInfo.area_hall,2)} - {round(FloorInfo.area_floorRemoveCar*0.1,2)} = {round(FloorInfo.over_hall,2)}...計入容積")
                addTextLine(msp, FloorInfo, content, 2)

        if FloorInfo.area_committee >0.:
            content = str(f"管委會面積: {round(FloorInfo.area_committee,2)}")
            addTextLine(msp, FloorInfo, content, 2)

        if FloorInfo.area_electromechanical >0.:
            content = str(f"機電面積: {round(FloorInfo.area_electromechanical,2)}")
            addTextLine(msp, FloorInfo, content, 2)

        if FloorInfo.area_balcony >0.:
            content = str(f"陽台面積: {round(FloorInfo.area_balcony,2)}")
            addTextLine(msp, FloorInfo, content, 1)
            
            if FloorInfo.over_balcony<0.001:
                content = str(f"{round(FloorInfo.area_balcony,2)} < {round(FloorInfo.area_floorRemoveCar,2)}*0.1 = {round(FloorInfo.area_floorRemoveCar*0.1,2)}...OK")
                addTextLine(msp, FloorInfo, content, 2)
                
            else:
                content = str(f"{round(FloorInfo.area_balcony,2)} > {round(FloorInfo.area_floorRemoveCar,2)}*0.1 = {round(FloorInfo.area_floorRemoveCar*0.1,2)}")
                addTextLine(msp, FloorInfo, content, 1)
                
                content = str(f"{round(FloorInfo.area_balcony,2)} - {round(FloorInfo.area_floorRemoveCar*0.1,2)} = {round(FloorInfo.over_balcony,2)}...計入容積")
                addTextLine(msp, FloorInfo, content, 2)

        content = str("梯廳+陽台面積檢討")
        addTextLine(msp, FloorInfo, content, 1)

        content = str(f"{round(FloorInfo.area_hall,2)}+{round(FloorInfo.area_balcony,2)}={round(FloorInfo.area_hall,2)+round(FloorInfo.area_balcony,2)}")
        addTextLine(msp, FloorInfo, content, 1)
        
        if (FloorInfo.area_hall + FloorInfo.area_balcony) < (FloorInfo.area_floorRemoveCar)*0.15:
            content = str(f"{round(FloorInfo.area_hall,2)+round(FloorInfo.area_balcony,2)} < {round(FloorInfo.area_floorRemoveCar,2)}*0.15= {round(FloorInfo.area_floorRemoveCar*0.15,2)} ...OK")
            addTextLine(msp, FloorInfo, content, 2)
        else:
            content = str(f"{round(FloorInfo.area_hall,2)+round(FloorInfo.area_balcony,2)} > {round(FloorInfo.area_floorRemoveCar,2)}*0.15= {round(FloorInfo.area_floorRemoveCar*0.15,2)}")
            addTextLine(msp, FloorInfo, content, 1)

            content = str(f"{round(FloorInfo.area_hall,2)+round(FloorInfo.area_balcony,2)} - {round(FloorInfo.area_floorRemoveCar*0.15,2)} - {round(FloorInfo.over_hall,2)}(10%梯廳超出) - {round(FloorInfo.over_balcony,2)}(10%陽台超出) = {round(FloorInfo.over_hallBalcony,2)}...計入容積")
            addTextLine(msp, FloorInfo, content, 2)

        content = str("容積樓地板面積")
        addTextLine(msp, FloorInfo, content, 1)

        content = str(f"= {round(FloorInfo.area_floor,2)}(樓地板面積) - {round(FloorInfo.area_hall,2)}(梯廳面積) - {round(FloorInfo.area_committee,2)}(管委會面積) - {round(FloorInfo.area_electromechanical,2)}(機電面積) - {round(FloorInfo.area_carramp,2)}(車道面積) + {round(FloorInfo.over_hall,2)}(10%梯廳超出) + {round(FloorInfo.over_balcony,2)}(10%陽台超出) + {round(FloorInfo.over_hallBalcony,2)}(15%梯廳陽台超出)")
        addTextLine(msp, FloorInfo, content, 1)

        content = str(f"= {round(FloorInfo.area_volumn,2)}")
        addTextLine(msp, FloorInfo, content, 1)