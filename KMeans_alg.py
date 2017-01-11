import sys, time, datetime
import random, math
import win32com.client, MSO, MSPPT

from MSO import constants as MSOcon
from MSPPT import constants as MSPPTcon

import KMeans_defs
from KMeans_defs import Sample
from KMeans_defs import Mean
from KMeans_defs import Cluster
from KMeans_defs import KMeansParameters

tmp_save_path = "C:\\Dev\\PyToPPT\\"
now_string = str.replace(datetime.datetime.now().isoformat(), ':', '.')
tmp_save_filename = tmp_save_path + "test-" + now_string + ".pptx"
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Add()
Presentation.saveas(tmp_save_filename)


#===== init =======
def initialize_alg_parameters():
    KMeans_defs.Base = Presentation.Slides.Add(1, 12)
    alg_parameters = KMeansParameters()
    clusters = alg_parameters.clusters

    KMeans_defs.create_axis(275, 50, 275, 500)
    KMeans_defs.create_axis(30, 275, 550, 275)

    samples = []
    total_samples = KMeans_defs.Total_Samples
    for i in range(total_samples):
        cluster_index = i/(total_samples/3)
        rand_left_coord = random.randint(clusters[cluster_index].left_start, clusters[cluster_index].left_end)
        rand_top_coord = random.randint(clusters[cluster_index].top_start, clusters[cluster_index].top_end)
        samples.append(Sample(rand_left_coord, rand_top_coord))

    return samples

#===== K-Means =======
def KMeans_alg():
    samples = initialize_alg_parameters()

    # 1. Randomly select centers
    means = []
    total_samples = len(samples)
    rand_means = random.sample(range(0, total_samples), KMeans_defs.Number_of_Clusters)
    i = 0
    for cluster_type in KMeans_defs.clusters_dict.iteritems():
        random_sample_index = rand_means[i]
        new_mean = Mean(samples[random_sample_index].left(), samples[random_sample_index].top(), cluster_type[0])
        means.append(new_mean)
        i += 1

    for alg_iteration in range(KMeans_defs.Total_Algorithm_Iterations):
        #2. For each sample, for each mean: calculate distance, asign mean
        for curr_sample in samples:
            min_dist = float("inf")
            for curr_mean in means:
                dist_from_curr_mean = curr_mean.dist_from_sample(curr_sample)
                if dist_from_curr_mean < min_dist:
                    min_dist = dist_from_curr_mean
                    curr_sample.classify(curr_mean)

        time.sleep(1)

        #3. For each mean, calculate its new coords
        for curr_mean in means:
            sum_left_coord = 0
            sum_top_coord = 0
            samples_belong_to_mean = 0
            for curr_sample in samples:
                if curr_sample.get_class_name() == curr_mean.class_name:
                    sum_left_coord += curr_sample.left()
                    sum_top_coord += curr_sample.top()
                    samples_belong_to_mean += 1

            curr_mean.left = sum_left_coord/samples_belong_to_mean
            curr_mean.top = sum_top_coord/samples_belong_to_mean
            curr_mean.replace_mean_shape()

        time.sleep(1)

    for i in range(len(means)):
        print(means[i].left, means[i].top)

    Presentation.save()


#===== trial-and-error =======
#def animate(seq, shape, trigger, path, duration=1.5):

def animate(seq, shape, trigger, path, duration=1.5):
    '''Move image along path when trigger is clicked'''


def animate_try():
    KMeans_defs.Base = Presentation.Slides.Add(1, 12)
    shape = KMeans_defs.Base.Shapes.AddShape(MSOcon.msoShapeOval, 100, 100, 50, 50)

    #motion_effect = Presentation.Slides(1).TimeLine.MainSequence.Behaviors.Add(MSOcon.msoAnimTypeMotion).MotionEffect
    seq = Presentation.Slides(1).TimeLine.MainSequence
    motion_effect_def = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(shape, effectId=MSPPTcon.msoAnimTypeMotion)
    effect = motion_effect_def.behaviors.add(MSPPTcon.msoAnimTypeMotion)
    effect.motioneffect.fromX = 0
    effect.motioneffect.fromY = 0
    effect.motioneffect.toX = 20
    effect.motioneffect.toY = 50
    motion_effect_def.Timing.Duration = 1

    shape2 = KMeans_defs.Base.Shapes.AddShape(MSOcon.msoShapeOval, 300, 100, 80, 80)
    motion_effect_def = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(shape2,
                                                                               effectId=MSPPTcon.msoAnimTypeMotion,
                                                                               trigger=MSPPTcon.msoAnimTriggerWithPrevious)
    effect = motion_effect_def.behaviors.add(MSPPTcon.msoAnimTypeMotion)
    effect.motioneffect.fromX = 0
    effect.motioneffect.fromY = 0
    effect.motioneffect.toX = 20
    effect.motioneffect.toY = -20
    motion_effect_def.Timing.Duration = 1

    #shape2.left = 300 + 20.0*540./100
    #shape2.top = 300 - 20.0*540./100
    motion_effect_def2 = Presentation.Slides(1).TimeLine.MainSequence.AddEffect(shape2,
                                                                               effectId=MSPPTcon.msoAnimTypeMotion,
                                                                               trigger=MSPPTcon.msoAnimTriggerOnPageClick)

    effect2 = motion_effect_def2.behaviors.add(MSPPTcon.msoAnimTypeMotion)
    effect2.motioneffect.fromX = effect.motioneffect.toX
    effect2.motioneffect.fromY = effect.motioneffect.toY
    effect2.motioneffect.toX = effect.motioneffect.fromX
    effect2.motioneffect.toY = effect.motioneffect.fromY
    motion_effect_def2.Timing.Duration = 2

    Presentation.save()
    #animate()

#========= Execute =========
animate_try()

