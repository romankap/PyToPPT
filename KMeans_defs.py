import sys, win32com.client, MSO, MSPPT, time, random, math
g = globals()
random.seed()

from MSO import constants as MSOcon
from MSPPT import constants as MSPPTcon
global Presentation
global Base
global Slide_Width
global Slide_Height

#for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
#for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)


#--------- Definitions ----------
def RGBtoInt(r, g, b):
    return r + g*256 + b*256*256

yellow_color = RGBtoInt(255, 255, 0)
blue_color = RGBtoInt(0, 0, 255)
red_color = RGBtoInt(255, 0, 0)
black_color = RGBtoInt(0,0,0)

def convert_pixels_to_screen_percentage(old_left, new_left, old_top, new_top):
    delta_X = 100.*(new_left - old_left)/Slide_Width
    delta_Y = 100.*(new_top - old_top)/Slide_Height
    return (delta_X, delta_Y)

def animate(seq, shape, trigger, path, duration=1.5):
    '''Move image along path when trigger is clicked'''
    effect = seq.AddEffect(
       Shape=shape,
       effectId=MSOcon.msoAnimEffectPathDown,
       trigger=MSOcon.msoAnimTriggerOnShapeClick,
    )
    ani = effect.Behaviors.Add(MSOcon.msoAnimTypeMotion)
    ani.MotionEffect.Path = path
    effect.Timing.TriggerType = MSOcon.msoAnimTriggerWithPrevious
    effect.Timing.TriggerShape = trigger
    effect.Timing.Duration = duration


def create_axis(left_start, top_start, left_end, top_end):
    axis = Base.Shapes.AddLine(left_start, top_start, left_end, top_end)
    axis.line.foreColor.RGB = 0
    axis.line.weight = 3.5
    axis.line.EndArrowheadStyle = MSOcon.msoArrowheadTriangle

#--------- Definitions ----------
Total_Algorithm_Iterations = 5
Total_Samples = 90
Number_of_Clusters = 3
global clusters_dict
clusters_dict = {"yellow": yellow_color, "blue": blue_color, "red": red_color}

cluster1_left_start = 50;    cluster1_left_stop=200
cluster1_top_start = 50;    cluster1_top_stop=230

cluster2_left_start = 220;   cluster2_left_stop=450
cluster2_top_start = 50;     cluster2_top_stop=450

cluster3_left_start = 50;   cluster3_left_stop=230
cluster3_top_start = 270;   cluster3_top_stop=450

Sample_Size = 40
Mean_Size = 40

#------- Classes ---------
class Sample:
    def __init__(self, left_coord, top_coord, color=black_color):
        self.shape = Base.Shapes.AddShape(MSOcon.msoShapeOval, left_coord, top_coord, Sample_Size, Sample_Size)
        self.shape.fill.forecolor.RGB = color
        self.shape.Line.ForeColor.RGB = black_color
        self.shape.Line.Weight = 1
        #self.shape.TextFrame.TextRange.Text = str(top_coord)
        #self.shape.TextFrame.TextRange.font.size = 12
        self.class_name = None
        self.color = black_color
        self.left = left_coord
        self.top = top_coord
        self.effects = []

    def left(self):
        return self.left

    def top(self):
        return self.top

    def set_fill_color(self, new_color):
        self.shape.fill.forecolor.RGB = new_color

    def get_class_name(self):
        return self.class_name

    def classify(self, mean):
        self.class_name = mean.class_name
        self.color = mean.color
        self.set_fill_color(mean.color)

    def add_classification_animation(self, mean, is_first=False):
        if is_first:
            trigger = MSPPTcon.msoAnimTriggerOnPageClick
        else:
            trigger = MSPPTcon.msoAnimTriggerWithPrevious
        color_effect_def = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(self.shape,
                              effectId=MSPPTcon.msoAnimEffectChangeFillColor, trigger=trigger)


        color_effect = color_effect_def.behaviors.add(MSPPTcon.msoAnimTypeColor)
        color_effect.ColorEffect.From.RGB = self.color
        color_effect.ColorEffect.To.RGB = mean.color
        color_effect_def.Timing.duration = 1

        self.class_name = mean.class_name
        self.color = mean.color
        self.effects.append(color_effect_def)


class Mean:
    def __init__(self, _left, _top, _class):
        self.left = _left
        self.top = _top
        self.class_name = _class
        self.color = clusters_dict[self.class_name]
        self.shape = None
        self.create_mean_shape()
        self.animation_effect = None
        self.animation_effect_new_left = 0
        self.animation_effect_new_top = 0

    def create_mean_shape(self):
        self.shape = Base.Shapes.AddShape(MSOcon.msoShapeMathMultiply, self.left, self.top, Mean_Size, Mean_Size)
        self.shape.line.weight = 3
        self.shape.line.forecolor.RGB = black_color
        self.shape.fill.forecolor.RGB = self.color
        #self.shape.line.forecolor.RGB = 1

    def replace_mean_shape(self):
        self.shape.delete()
        self.create_mean_shape()

    def dist_from_sample(self, sample):
        return math.pow(self.left - sample.left, 2) + math.pow(self.top - sample.top, 2)

    def add_motion_animation(self, new_left, new_top, is_first=False):
        if is_first:
            #seq = Presentation.Slides(1).TimeLine.MainSequence

            motion_effect_def = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(self.shape,
                                    effectId=MSPPTcon.msoAnimTypeMotion, trigger=MSPPTcon.msoAnimTriggerOnPageClick)
        else:
            motion_effect_def = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(self.shape,
                                    effectId=MSPPTcon.msoAnimTypeMotion, trigger=MSPPTcon.msoAnimTriggerWithPrevious)
        effect = motion_effect_def.behaviors.add(MSPPTcon.msoAnimTypeMotion)
        if not self.animation_effect:
            effect.motioneffect.fromX = 0
            effect.motioneffect.fromY = 0
        else:
            effect.motioneffect.fromX = self.animation_effect.toX
            effect.motioneffect.fromY = self.animation_effect.toY

        dest_coords = convert_pixels_to_screen_percentage(self.shape.left, new_left, self.shape.top, new_top)
        effect.motioneffect.toX = dest_coords[0]
        effect.motioneffect.toY = dest_coords[1]
        motion_effect_def.Timing.Duration = 1
        self.left = new_left
        self.top = new_top
        self.animation_effect = effect.motioneffect

class Cluster:
    def __init__(self, _left_start, _left_end, _top_start, _top_end):
        self.left_start =   _left_start
        self.left_end =     _left_end
        self.top_start =    _top_start
        self.top_end =      _top_end

class KMeansParameters:
    def __init__(self):
        self.clusters = []
        self.clusters.append(Cluster(cluster1_left_start, cluster1_left_stop, cluster1_top_start, cluster1_top_stop))
        self.clusters.append(Cluster(cluster2_left_start, cluster2_left_stop, cluster2_top_start, cluster2_top_stop))
        self.clusters.append(Cluster(cluster3_left_start, cluster3_left_stop, cluster3_top_start, cluster3_top_stop))
