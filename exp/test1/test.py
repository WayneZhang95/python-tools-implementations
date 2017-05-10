#!/usr/bin/env python
# -*- coding: utf-8 -*-
import json
from datetime import datetime


if __name__ == '__main__':
    with open('%s.txt' % datetime.now().strftime('%H-%M-%S'), 'w') as f:
        f.write('test1' + json.dumps(datetime.now().strftime('%Y-%m-%d-%H:%M:%S')))


