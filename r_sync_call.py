import shutil
import glob
import os
import time
import os.path as osp
from datetime import datetime

import clarity

from util import path
from shot.workflow.util import pathfunc
# from shot.workflow.util.osutil import copy_file


# PATH = os.environ["MAYACONTENT"]
OUT_GOING = r"OUTGOING_DIR"

os.environ["OUTGOING_DIR"] = r"\\10.0.132.38\Rsync"


def copy_file(src, dst):
    if not osp.isdir(osp.dirname(dst)):
        os.makedirs(osp.dirname(dst))

    shutil.copyfile(src, dst)


def _copy_movie(shot_code, step):

    mov_file, first_file, second_file, third_file = path.get_mov_client(shot_code, step)

    copy_file(mov_file, first_file)
    copy_file(mov_file, second_file)
    copy_file(mov_file, third_file)

    return mov_file, first_file, second_file, third_file


def _copy_maya(shot_code, step):
    import preprocess

    msg = None
    _, scene_file = pathfunc.get_latest_scene_task(pathfunc.IS_REVIEW, shot_code, b_steps=[step],
                                                   exact_step=True)
    if not osp.isfile(scene_file):
        raise IOError("Please check shot node again, Cannot find it in server's folder")

    dst_file = path.get_maya_client(shot_code, step, scene_file)

    if osp.isfile(dst_file):
        return True, dst_file

    temp_file = path.get_maya_temp(scene_file)
    copy_file(scene_file, temp_file)

    try:
        preprocess.pre_process(temp_file,dst_file)
    except IOError as e:
        msg = "Cannot pre_process " + str(e)

    return msg, dst_file


def get_shotcode():
    # eps = "312A"
    EXCEPTION_DCT = {
        "eps": ["306a", "999A", "EPI"],
        "seq": ["master"]
    }
    lst_shotcode = []
    SHOT_REVIEW = os.environ["REV_SHOT_CENTRAL"]
    for eps in sorted(os.listdir(SHOT_REVIEW)):
        if eps not in EXCEPTION_DCT["eps"]:
            eps_path = osp.join(SHOT_REVIEW,eps)
            for seq in sorted(os.listdir(eps_path)):
                if seq not in EXCEPTION_DCT["seq"]:
                    seq_path = osp.join(eps_path, seq)
                    for shot in sorted(os.listdir(seq_path)):
                        shot_code = ".".join([eps,seq, shot])
                        lst_shotcode.append(shot_code)

    return lst_shotcode


def run(*args):
    print "\nR-Sync Tool\n"
    start_time = time.time()
    lst_error = []
    lst_dont_update = []
    lst_success = []

    lst_shotcode = get_shotcode()
    lst_step = args[0]
    print lst_step
    for shot_code in lst_shotcode:
        for step in lst_step:
            shot_step = ".".join([shot_code,step])
            print shot_step
            try:
                msg, scene_maya = _copy_maya(shot_code, step)
                if type(msg) is str:
                    raise IOError(msg)
                elif type(msg) is bool and msg:
                    lst_dont_update.append(shot_step)
                    continue
                mov_file, first_file, second_file, third_file = _copy_movie(shot_code, step)
                lst_success.append([shot_step, osp.basename(mov_file), first_file, second_file, third_file])

            except Exception as e:
                lst_error.append([shot_step, str(e)])

    LOG_FILE = osp.join(os.environ["LOG_EVENT"], "r_sync", "delivery_{date}.txt")
    log_file = LOG_FILE.format(date=datetime.now().strftime("%Y%m%d_%H%M%S"))
    if not osp.isdir(osp.dirname(log_file)):
        os.makedirs(osp.dirname(log_file))
    with open(log_file, "w") as f:
        f.write("It's take {} minutes".format((time.time()-start_time)/60))
        for i, e in lst_error:
            f.write("Error shot {}\n".format(i))
            f.write(e + "\n")
            f.write("End Error shot {}\n".format(i))
            f.write("\n")
        for shot_code, i, first, s, t in lst_success:
            f.write("Success shot {}\n".format(shot_code))
            f.write(i + "\n")
            f.write(first + "\n")
            f.write(s + "\n")
            f.write(t + "\n")
            f.write("End Success shot {}\n".format(shot_code))
            f.write("\n")

        if lst_dont_update:
            f.write("Don't need to update shot, because we have the lastest version in R-Sync \n")
            for shot_code in lst_dont_update:
                f.write(shot_code + "\n")
            f.write("End list shot\n")

    # remove temp:
    if osp.isdir(path.FOLDER_TEMP):
        shutil.rmtree(path.FOLDER_TEMP)

    # os.system(log_file)
    # raw_input("Please press enter to exit.")


if __name__ == "__main__":
    import sys
    run(sys.argv[1:])
