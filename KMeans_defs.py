import sys, win32com.client, MSO, MSPPT, time, random
g = globals()
random.seed()

from MSO import constants as MSOcon

#for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
#for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)


#--------- Definitions ----------
def RGBtoInt(r, g, b):
    return r + g*256 + b*256*256

yellow_color = RGBtoInt(255, 255, 0)
blue_color = RGBtoInt(0, 0, 255)
red_color = RGBtoInt(255, 0, 0)
black_color = RGBtoInt(0,0,0)

def animate(seq, image, trigger, path, duration=1.5):
    '''Move image along path when trigger is clicked'''
    effect = seq.AddEffect(
       Shape=image,
       effectId=MSO.constants.msoAnimEffectPathDown,
       trigger=MSO.constants.msoAnimTriggerOnShapeClick,
    )
    ani = effect.Behaviors.Add(MSO.constants.msoAnimTypeMotion)
    ani.MotionEffect.Path = path
    effect.Timing.TriggerType = MSO.constants.msoAnimTriggerWithPrevious
    effect.Timing.TriggerShape = trigger
    effect.Timing.Duration = duration

def create_new_sample(left_coord, top_coord, color=black_color):
    sample = Base.Shapes.AddShape(MSOcon.msoShapeOval, left_coord, top_coord, 40, 40)
    sample.fill.forecolor.RGB = color
    sample.Line.ForeColor.RGB = black_color
    sample.Line.Weight = 1
    sample.TextFrame.TextRange.Text = str(top_coord)
    sample.TextFrame.TextRange.font.size = 12

    return sample

def create_axis(left_start, top_start, left_end, top_end):
    axis = Base.Shapes.AddLine(left_start, top_start, left_end, top_end)
    axis.line.foreColor.RGB = 0
    axis.line.weight = 3.5
    axis.line.EndArrowheadStyle = MSO.constants.msoArrowheadTriangle

#--------- Definitions ----------
number_of_clusters = 3
clusters_dict = {"yellow": yellow_color, "blue": blue_color, "red": red_color}

cluster1_x_start = 50;    cluster1_x_stop=200
cluster1_y_start = 50;    cluster1_y_stop=200

cluster2_x_start = 300;   cluster2_x_stop=500
cluster2_y_start = 50;    cluster2_y_stop=200

cluster3_x_start = 100;   cluster3_x_stop=350
cluster3_y_start = 300;   cluster3_y_stop=450

class Mean:
    def __init__(self, _left, _top, _class):
        self.left = _left
        self.top = _top
        self.class_name = _class
        self.color = clusters_dict[self.class_name]
        self.shape = None
        self.create_mean_shape()

    def create_mean_shape(self):
        self.shape = Base.Shapes.AddShape(MSO.constants.msoShapeMathMultiply, self.left, self.top, 40, 40)
        self.shape.line.weight = 1
        self.shape.fill.forecolor.RGB = self.color
        #self.shape.line.forecolor.RGB = 1

    def dist_from_sample(self, sample):
        return (self.left - sample.left)**2 + (self.top - sample.top)**2

class Cluster:
    def __init__(self, _x_start, _x_end, _y_start, _y_end):
        self.x_start =  _x_start
        self.x_end =    _x_end
        self.y_start =  _y_start
        self.y_end =    _y_end

class KMeansParameters:
    def __init__(self):
        self.clusters = []
        self.clusters.append(Cluster(cluster1_x_start, cluster1_x_stop, cluster1_y_start, cluster1_y_stop))
        self.clusters.append(Cluster(cluster2_x_start, cluster2_x_stop, cluster2_y_start, cluster2_y_stop))
        self.clusters.append(Cluster(cluster3_x_start, cluster3_x_stop, cluster3_y_start, cluster3_y_stop))


#----- Slide Generation

# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Add()
Base = Presentation.Slides.Add(1, 12)

samples = []
total_samples = 12

alg_parameters = KMeansParameters()
clusters = alg_parameters.clusters

for i in range(total_samples):
    cluster_index = i/(total_samples/3)
    rand_x_coord = random.randint(clusters[cluster_index].x_start, clusters[cluster_index].x_end)
    rand_y_coord = random.randint(clusters[cluster_index].y_start, clusters[cluster_index].y_end)
    samples.append(create_new_sample(rand_x_coord, rand_y_coord))

create_axis(275, 50, 275, 500)
create_axis(30, 275, 550, 275)

#===== K-Means =======
# 1. Randomly select centers
centers = []
rand_means = random.sample(range(0, total_samples), number_of_clusters)
i = 0
for cluster_type in clusters_dict.iteritems():
    random_sample_index = rand_means[i]
    new_mean = Mean(samples[random_sample_index].left, samples[random_sample_index].top, cluster_type[0])
    centers.append(new_mean)
    i += 1

for i in range(len(centers)):
    print(centers[i].left, centers[i].top)

    # Add an oval. Shape 9 is an oval.
#oval = Base.Shapes.AddShape(9, 2, 2, 2, 2)
#oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
#line = Base.Shapes.AddShape(0x6d, 0, 100, 100, 100)
#oval2 = Base.Shapes.AddShape(9, 100, 0, 50, 50)
#oval.Fill.ForeColor.RGB = RGB(255, 0, 0)
#oval.top = 200


#for i in range(1, 50):
#    Base.Shapes.AddShape(i, 100, i*5, 5, 5)
#line = Base.Shapes.AddLine(10, 300, 200, 300)
#line.line.foreColor.RGB = 0
#line.line.dashstyle = MSO.constants.msoLineThickThin

#oval = Base.Shapes.AddShape(9, 0, 100, 100, 100)
#oval.Fill.ForeColor.RGB = RGB(255, 0, 0)