import numpy as np

class CPoint():
    def __init__(self, x:float, y:float):
        self.x=x
        self.y=y

    def __str__(self):
        return f"x:{self.x}, y:{self.y}"

    def add(self, p2):
        return CPoint(self.x+p2.x, self.y+p2.y)

    def sub(self, p2):
        return CPoint(self.x-p2.x, self.y- p2.y)


class CLine():
    def __init__(self, p1:CPoint, p2:CPoint):
        self.p1=p1
        self.p2=p2

    def getLength(self):
        return  pow(self.p1.x - self.p2.x, 2 ) +pow(self) 

    def __str__(self):
        return f"Line: [{self.p1}, {self.p2}]"
    
    def is_horizon(self):
        x = round(abs(self.p1.x - self.p2.x),2)
        y = round(abs(self.p1.y - self.p2.y),2)
        if x >0 and y==0:
            return True

# p2---------p1
#  |         |
#  |         |
#  |         |
# p3---------p4
class CSqure():
    def __init__(self, line1:CLine, line2:CLine):
        self.l1=line1
        self.l2=line2
        self.p1 = line1.p1
        self.p2 = line1.p2
        self.p3 = line2.p2
        self.p4 = line2.p1
        self.points = [(self.p1.x,self.p1.y),(self.p2.x,self.p2.y),(self.p3.x,self.p3.y),(self.p4.x,self.p4.y),(self.p1.x,self.p1.y)]

    def get_4_point(self):
        return self.points
    
    def get_centor(self):
        centor_x = (self.p1.x+self.p3.x)/2
        centor_y = (self.p1.y+self.p3.y)/2
        centor = CPoint(centor_x,centor_y)
        return centor
    
    def x_min(self):
        x_min = min(point[0] for point in self.points)
        return x_min
    
    def x_max(self):
        x_max = max(point[0] for point in self.points)
        return x_max
    
    def y_min(self):
        y_min = min(point[1] for point in self.points)
        return y_min
    
    def y_max(self):
        #y_max = max(self.points[:][1])
        y_max = max(point[1] for point in self.points)
        return y_max


class CPointVector():
    def __init__(self, p0=CPoint(0,0), p1=CPoint(1,1)):
        self.vx = p1.x-p0.x
        self.vy = p1.y-p0.y

    def __str__(self):
        return f"Vector: [{self.vx}, {self.vy}]"

class point_list():
    def __init__(self, LWP_point):
        self.LWP_point = LWP_point
    
    def result(self):
        points = []
        for inx, p in enumerate(self.LWP_point):
            points.append(CPoint(p[0],p[1]))
        return points

class floor_imformation():
    def __init__(self, x_min:float, x_max:float, y_min:float, y_max:float):
        self.x_min = x_min
        self.x_max = x_max
        self.y_min = y_min
        self.y_max = y_max
        self.write_line = 0
        self.font_size = abs(y_max-y_min)/50
        self.write_point = CPoint(self.x_max,self.y_max)

        self.area_floor = 0.00
        self.area_floorRemoveCar = 0.00
        self.area_hall = 0.00
        self.area_committee = 0.00
        self.area_electromechanical = 0.00
        self.area_carramp = 0.00
        self.area_balcony = 0.00
        self.area_volumn = 0.00

        self.over_hall =  0.00
        self.over_balcony =  0.00
        self.over_hallBalcony =  0.00

        self.name = "new floor"


class detA_sum():
    def __init__(self, point_num:int, point):
        self.point_num = point_num
        self.point = point
    
    def result(self):
        origin_vector = []
        sum_detA = 0.000
        for i in range(self.point_num):
            vector_i = CPointVector(self.point[i],self.point[0])
            origin_vector.append(vector_i)

        for i in range(len(origin_vector)-1):
            matrix = [[origin_vector[i].vx,origin_vector[i].vy],
                      [origin_vector[i+1].vx,origin_vector[i+1].vy]]
            #print(f"Matrix {i+1} = {matrix}")
            s= np.linalg.det(matrix)*0.5
            sum_detA += s
        return sum_detA