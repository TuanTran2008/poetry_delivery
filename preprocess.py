import os
import os.path as osp
import sys

import maya.cmds as mc
import maya.mel as mel
import pymel.core as pm
import rbwcgi_launcher_ui as ui
# sys.path.append(r"C:\vsTools2\SPD\python")
from shot.wrangler.anim import builder

SPARX_CAM = "CamStandard_001:ploader_001"
CLIENT_CAM = "viewtool_camera_df_:Render_Cam"

START_FRAME = 1001
CLT_START_FRAME = 001


def _is_lost_animtaion(new_scene):
    with open(new_scene) as f:
        contents = f.read()
    result = "file -r -ns" in contents and "createNode reference" not in contents
    return result


def _move_anim_curve(start_frame, start_shot, start_frame_shot_node):
    if start_frame == start_shot and start_frame_shot_node == start_frame:
        return True
    lst_curve = [x for x in mc.ls(type="animCurve")]
    len_lst = len(lst_curve)
    for idx, curve in enumerate(lst_curve):
        print "Moving Anim Curve {}%".format(float(idx + 1) / len_lst * 100)
        start_curve = (mc.keyframe(curve, query=True, tc=True) or [None])[0]
        if start_curve and start_curve >= start_frame:
            try:
                mc.keyframe(curve, edit=True, iub=False, r=True, o="over", tc=start_frame - start_shot)
            except Exception as e:
                if str(e).strip() == "Cannot move keys":
                    mc.setAttr(curve + ".ktv", lock=True)
                    mc.setAttr(curve + ".ktv", lock=False)
                    mc.keyframe(curve, edit=True, iub=False, r=True, o="over", tc=start_frame - start_shot)

    duration = int(mc.playbackOptions(query=True, aet=True)) - start_shot

    mc.playbackOptions(ast=start_frame, minTime=start_frame, aet=start_frame + duration, maxTime=start_frame + duration)


def _unlock_camera_attr(camera_client):
    # Turn on force lockeditable when publish layout-scene
    mel.eval("optionVar -iv refLockEditable true")
    # lock camera's client
    dct_attr_lock = {
        "translate": False,
        "rotate": False,
        "scale": False,
        "hfa": False,
        "vfa": False,
        "lensSqueezeRatio": False,
        "ffo": False,
        "horizontalFilmOffset": False,
        "verticalFilmOffset": False,
        "ptsc": False,
        "frv": False,
        "horizontalRollPivot": False,
        "verticalRollPivot": False,
        "filmTranslateH": False,
        "filmTranslateV": False,
        "psc": False,
    }

    for attr in dct_attr_lock.keys():
        eval('camera_client.{}.set(l={})'.format(attr, dct_attr_lock[attr]))

    camera_client.overscan.set(1)
    camera_client.displaySafeAction.set(False)


def _remove_cam():
    cam_obj = SPARX_CAM

    if len(pm.ls(CLIENT_CAM)) == 0:
        raise IOError("Can't find camera of client. Please check a scene again")
    elif len(pm.ls(CLIENT_CAM)) > 1:
        raise IOError("The name {} isn't unique. Please check a scene again".format(CLIENT_CAM))

    cam_client = pm.PyNode(CLIENT_CAM)

    _unlock_camera_attr(cam_client)

    if pm.objExists(cam_obj):
        pm.select(cam_obj)
        builder.unload_asset(pm.selected(), 1, 1)
        pm.delete("cam_grp")


def remove_unknowplugin():
    lst_unknowplugin = pm.unknownPlugin(query=True, list=1) or []
    for plugin in lst_unknowplugin:
        try:
            pm.unknownPlugin(plugin, r=True)
        except RuntimeError:
            print plugin


def _force_load_all_reference():
    rf_nodes = pm.ls(references=True)
    for rfn in rf_nodes:
        try:
            if rfn.locked.get() or rfn.isLoaded():
                print "Reference {} is loaded".format(rfn.nodeName())
                continue
            rfnode = pm.FileReference(rfn)
            rfnode.load()
        except RuntimeError as ex:
            continue


# tag the remote pipeline release in the top group LIBRARY
def tag_remote_library():
    if pm.objExists('LIBRARY'):
        try:
            pm.addAttr('LIBRARY', longName='remote_pipeline_release', dt='string')
        except:
            pm.setAttr('LIBRARY' + '.remote_pipeline_release', l=0)
        pm.setAttr('LIBRARY.remote_pipeline_release', ui.remoteVersion, type="string")
        pm.setAttr('LIBRARY' + '.remote_pipeline_release', l=1)


def is_existed_anim_layer():
    lst_anim_layer = pm.ls(type="animLayer")[:-1]

    for anim_layer in lst_anim_layer:
        if pm.animLayer(anim_layer, q=1, anc=1):
            return True
    return False



def pre_process(scene_file, dst_file=None):
    try:
        project_dir = os.environ['MAYACONTENT']
        pm.workspace(project_dir, openWorkspace=True)
        print "workspace is seted", project_dir
        # openScene()
        pm.openFile(scene_file, force=True)

        if is_existed_anim_layer():
            raise IOError("Please Remove Anim Layer before delivery")

        print "Trying load all reference of {}".format(osp.basename(scene_file))
        _force_load_all_reference()
        print "Finish load all reference"

        # Remove Camera:
        print "Trying delete shot node"
        start_frame_shot_node = delete_shot_node()
        print "Finish delete shot node"

        # Move anim_curve
        print "Trying moving curve"
        start_shot = int(pm.playbackOptions(query=True, ast=True))
        start_frame = int(os.environ["START_FRAME"])
        _move_anim_curve(start_frame, start_shot, start_frame_shot_node)
        print "Finish moving curve"

        # Modify sound():
        print "Trying delete cam of {}".format(osp.basename(scene_file))
        _remove_cam()
        print "Finish delete cam"

        # Remove Unknown Plugin
        print "Removing Unknown Plugin "
        remove_unknowplugin()
        print "Finish remove Unknown Plugin "

        print "Tagging extra attribute in LIBRARY group"
        tag_remote_library()
        print "Finish tag extra attribute in LIBRARY group "

        if dst_file is None:
            pm.saveFile()
            pm.newFile(f=1)
        else:
            # Make sure It's a correct path
            try:

                pm.saveFile()
                if _is_lost_animtaion(scene_file):
                    raise IOError("Have a trouble about connection with server, please deliver this shot again.")

                if not osp.isdir(osp.dirname(dst_file)):
                    os.makedirs(osp.dirname(dst_file))

                pm.saveAs(dst_file)
                if _is_lost_animtaion(dst_file):
                    raise IOError("Have a trouble about connection with server, please deliver this shot again.")
            except IOError as e:
                raise IOError(
                    "Can't save with this name {}, because can't connect to server, details:{}, please try again that shot".format(
                        dst_file, str(e)))

    except Exception as e:
        raise IOError(e)


def delete_shot_node():
    shot_node = ([x for x in pm.ls(type="shot") if x.startswith("Q")] + ["None"])[0]
    startframe = -1
    if pm.objExists(shot_node):
        startframe = int(shot_node.startFrame.get())
        pm.delete(shot_node)

    return startframe


if __name__ == "__main__":
    if sys.argv[1]:
        pre_process(sys.argv[1])
    else:
        print("Please input scene")
