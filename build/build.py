#!/usr/bin/env python
# -*- coding: utf-8-unix -*-

import base64
import codecs
import os

build = ""

with codecs.open('template.py', 'r', 'utf_8') as f:
    build += f.read()

with codecs.open('main.js', 'r', 'shift_jis') as f:
    build = build.replace("%%%GENERATE_FEATURE_JS%%%", f.read())

with codecs.open('launch.py', 'r', 'utf_8') as f:
    build += f.read()

target = '../generatefeature.py'
with codecs.open(target, 'w', 'utf_8') as f:
    f.write(build.replace("\r\n", "\n"))

os.chmod(target, 0755)
