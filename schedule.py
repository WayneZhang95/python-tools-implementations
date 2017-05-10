#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""* * * * * cd /home/jd/zz-scratch/spider && python /home/jd/zz-scratch/spider/test.py  // 已经cd到路径下了就不用再写绝对路径了

* * * * * cd /home/jd/zz-scratch/spider && python test.py  // 直接这样就好

中间脚本调用 """
import os
from glob import glob
import subprocess


def traverse_and_run():
    for path in glob('exp/*/'):
        for file_name in os.listdir(path):
            if file_name.endswith('.py'):
                p = subprocess.Popen(['python', file_name], cwd='%s' % path)
                p.wait()


if __name__ == '__main__':
    traverse_and_run()